/**
 * Minimal Work IQ Agent-to-Agent (A2A) client — JSON-RPC 2.0 over HTTPS.
 *
 * Work IQ exposes Microsoft 365 Copilot as an A2A agent: discover it via its
 * Agent Card, then exchange structured *tasks* (SendMessage / SendStreamingMessage
 * / GetTask / CancelTask) instead of calling a bespoke chat API. Uses the same
 * delegated token (scope WorkIQAgent.Ask) as the REST and MCP clients.
 *
 * Protocol version is selected with the `A2A-Version` header. Work IQ supports
 * 1.0 (default here) and 0.3 (the server default when the header is omitted).
 * Responses from either version are normalized to the 1.0 shapes below.
 *
 * Docs:
 *  - https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/a2a/overview
 *  - https://a2a-protocol.org/latest/specification/
 */

import { randomUUID } from "node:crypto";

const DEFAULT_BASE = "https://workiq.svc.cloud.microsoft/a2a";

export type A2aVersion = "1.0" | "0.3";

export interface A2aPart {
  text?: string;
  [key: string]: unknown;
}

export interface A2aMessage {
  role: string;
  messageId: string;
  contextId?: string;
  taskId?: string;
  parts: A2aPart[];
  metadata?: Record<string, unknown>;
}

export interface A2aArtifact {
  artifactId: string;
  name?: string;
  parts: A2aPart[];
}

export interface A2aTaskStatus {
  state: string;
  message?: A2aMessage;
  timestamp?: string;
}

export interface A2aTask {
  id: string;
  contextId: string;
  status: A2aTaskStatus;
  artifacts?: A2aArtifact[];
  history?: A2aMessage[];
}

export interface A2aStatusUpdate {
  taskId: string;
  contextId: string;
  status: A2aTaskStatus;
  final?: boolean;
}

export interface A2aArtifactUpdate {
  taskId: string;
  contextId: string;
  artifact: A2aArtifact;
  append?: boolean;
  lastChunk?: boolean;
}

/** A2A 1.0 StreamResponse / SendMessageResponse — exactly one field is set. */
export interface A2aStreamResponse {
  task?: A2aTask;
  message?: A2aMessage;
  statusUpdate?: A2aStatusUpdate;
  artifactUpdate?: A2aArtifactUpdate;
}

export interface A2aAgentCard {
  name: string;
  description?: string;
  version?: string;
  url?: string;
  supportedInterfaces?: Array<{
    url: string;
    protocolBinding?: string;
    protocolVersion?: string;
  }>;
  capabilities?: { streaming?: boolean; pushNotifications?: boolean };
  skills?: Array<{ id?: string; name?: string; description?: string }>;
  defaultInputModes?: string[];
  defaultOutputModes?: string[];
  [key: string]: unknown;
}

export interface A2aSendOptions {
  /** Continue a conversation: the `contextId` returned by a previous turn. */
  contextId?: string;
  /** Continue a specific task (e.g. one in TASK_STATE_INPUT_REQUIRED). */
  taskId?: string;
  /** IANA time zone sent as message metadata (Location). */
  timeZone?: string;
}

export class A2aRpcError extends Error {
  constructor(
    readonly method: string,
    readonly code: number,
    message: string,
    readonly data?: unknown
  ) {
    super(`Work IQ A2A ${method} error ${code}: ${message}`);
  }
}

/** v1.0 method name → v0.3 method name. */
const V03_METHODS: Record<string, string> = {
  SendMessage: "message/send",
  SendStreamingMessage: "message/stream",
  GetTask: "tasks/get",
  CancelTask: "tasks/cancel",
  SubscribeToTask: "tasks/resubscribe",
};

/** v0.3 lower-kebab task states → v1.0 enum names. */
function normalizeState(state: string | undefined): string {
  if (!state) return "TASK_STATE_UNSPECIFIED";
  if (state.startsWith("TASK_STATE_")) return state;
  return "TASK_STATE_" + state.replace(/-/g, "_").toUpperCase();
}

/** Minutes east of UTC for an IANA zone right now (e.g. Asia/Taipei → 480). */
export function timeZoneOffsetMinutes(timeZone: string): number {
  const now = new Date();
  const asUtc = (tz: string) =>
    new Date(now.toLocaleString("en-US", { timeZone: tz })).getTime();
  return Math.round((asUtc(timeZone) - asUtc("UTC")) / 60000);
}

export class WorkIqA2aClient {
  readonly baseUrl: string;
  private rpcUrl: string;
  private nextId = 1;

  constructor(
    private readonly getToken: () => Promise<string>,
    readonly options: {
      /** Target a specific Work IQ agent instead of the default Copilot agent. */
      agentId?: string;
      version?: A2aVersion;
      baseUrl?: string;
    } = {}
  ) {
    const base = (options.baseUrl ?? process.env.WORK_IQ_A2A_ENDPOINT ?? DEFAULT_BASE)
      .replace(/\/+$/, "");
    this.baseUrl = options.agentId
      ? `${base}/${encodeURIComponent(options.agentId)}`
      : base;
    this.rpcUrl = `${this.baseUrl}/`;
  }

  get version(): A2aVersion {
    return this.options.version ?? "1.0";
  }

  /** The JSON-RPC endpoint currently used for task/message calls. */
  get endpoint(): string {
    return this.rpcUrl;
  }

  private async headers(accept: string): Promise<Record<string, string>> {
    return {
      Authorization: `Bearer ${await this.getToken()}`,
      "Content-Type": "application/json",
      Accept: accept,
      "A2A-Version": this.version,
    };
  }

  private method(name: string): string {
    return this.version === "1.0" ? name : V03_METHODS[name] ?? name;
  }

  // ---- Discovery ----------------------------------------------------------

  /**
   * Fetch the Agent Card from `/.well-known/agent-card.json`. If it advertises a
   * JSON-RPC interface, later calls go to that URL (standard A2A discovery).
   */
  async getAgentCard(): Promise<A2aAgentCard> {
    const url = `${this.baseUrl}/.well-known/agent-card.json`;
    const res = await fetch(url, { headers: await this.headers("application/json") });
    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      throw new Error(`Work IQ A2A agent card failed: ${res.status} ${res.statusText}\n${detail}`);
    }
    const card = (await res.json()) as A2aAgentCard;
    const iface = card.supportedInterfaces?.find(
      (i) => (i.protocolBinding ?? "JSONRPC").toUpperCase() === "JSONRPC" && i.url
    );
    const advertised = iface?.url ?? card.url;
    if (advertised) this.rpcUrl = advertised.endsWith("/") ? advertised : `${advertised}/`;
    return card;
  }

  // ---- Messages -----------------------------------------------------------

  private buildMessage(text: string, opts: A2aSendOptions): Record<string, unknown> {
    const msg: Record<string, unknown> = {
      role: this.version === "1.0" ? "ROLE_USER" : "user",
      messageId: randomUUID(),
      parts: [this.version === "1.0" ? { text } : { kind: "text", text }],
    };
    if (this.version === "0.3") msg.kind = "message";
    if (opts.contextId) msg.contextId = opts.contextId;
    if (opts.taskId) msg.taskId = opts.taskId;
    if (opts.timeZone) {
      msg.metadata = {
        Location: {
          timeZone: opts.timeZone,
          timeZoneOffset: timeZoneOffsetMinutes(opts.timeZone),
        },
      };
    }
    return msg;
  }

  /** Low-level JSON-RPC call (non-streaming). Returns `result` or throws A2aRpcError. */
  async rpc<T = unknown>(method: string, params: unknown): Promise<T> {
    const wire = this.method(method);
    const res = await fetch(this.rpcUrl, {
      method: "POST",
      headers: await this.headers("application/json"),
      body: JSON.stringify({ jsonrpc: "2.0", id: this.nextId++, method: wire, params }),
    });
    const text = await res.text();
    let body: any = null;
    try {
      body = text ? JSON.parse(text) : null;
    } catch {
      /* non-JSON error page */
    }
    if (body?.error) {
      throw new A2aRpcError(wire, body.error.code, body.error.message, body.error.data);
    }
    if (!res.ok) {
      throw new Error(`Work IQ A2A ${wire} failed: ${res.status} ${res.statusText}\n${text}`);
    }
    return body?.result as T;
  }

  /** SendMessage — blocks until the task reaches a terminal / interrupted state. */
  async sendMessage(text: string, opts: A2aSendOptions = {}): Promise<A2aStreamResponse> {
    const result = await this.rpc("SendMessage", { message: this.buildMessage(text, opts) });
    return normalizeResponse(result);
  }

  /** SendStreamingMessage — yields each StreamResponse event from the SSE stream. */
  async *sendStreamingMessage(
    text: string,
    opts: A2aSendOptions = {}
  ): AsyncGenerator<A2aStreamResponse> {
    const wire = this.method("SendStreamingMessage");
    const res = await fetch(this.rpcUrl, {
      method: "POST",
      headers: await this.headers("text/event-stream"),
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: this.nextId++,
        method: wire,
        params: { message: this.buildMessage(text, opts) },
      }),
    });

    const contentType = res.headers.get("content-type") ?? "";
    if (!res.ok || !res.body || !contentType.includes("text/event-stream")) {
      const detail = await res.text().catch(() => "");
      let err: any = null;
      try {
        err = JSON.parse(detail)?.error;
      } catch {
        /* not JSON */
      }
      if (err) throw new A2aRpcError(wire, err.code, err.message, err.data);
      throw new Error(`Work IQ A2A ${wire} failed: ${res.status} ${res.statusText}\n${detail}`);
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let dataLines: string[] = [];

    const flush = (): A2aStreamResponse | null => {
      if (!dataLines.length) return null;
      const data = dataLines.join("\n");
      dataLines = [];
      let parsed: any;
      try {
        parsed = JSON.parse(data);
      } catch {
        return null; // keep-alive / non-JSON
      }
      if (parsed?.error) {
        throw new A2aRpcError(wire, parsed.error.code, parsed.error.message, parsed.error.data);
      }
      return parsed?.result ? normalizeResponse(parsed.result) : null;
    };

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      let idx: number;
      while ((idx = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, idx).replace(/\r$/, "");
        buffer = buffer.slice(idx + 1);
        if (line === "") {
          const evt = flush();
          if (evt) yield evt;
        } else if (line.startsWith("data:")) {
          dataLines.push(line.slice(5).trimStart());
        }
      }
    }
    if (buffer.startsWith("data:")) dataLines.push(buffer.slice(5).trimStart());
    const last = flush();
    if (last) yield last;
  }

  // ---- Tasks --------------------------------------------------------------

  async getTask(id: string, historyLength?: number): Promise<A2aTask> {
    const params: Record<string, unknown> = { id };
    if (historyLength !== undefined) params.historyLength = historyLength;
    const result = await this.rpc("GetTask", params);
    return normalizeTask(unwrapTask(result));
  }

  async cancelTask(id: string): Promise<A2aTask> {
    const result = await this.rpc("CancelTask", { id });
    return normalizeTask(unwrapTask(result));
  }
}

// ---- Normalization (v0.3 → v1.0 shapes) ------------------------------------

function unwrapTask(result: any): any {
  return result?.task ?? result;
}

function normalizeTask(t: any): A2aTask {
  return { ...t, status: { ...t?.status, state: normalizeState(t?.status?.state) } };
}

function normalizeResponse(result: any): A2aStreamResponse {
  if (!result || typeof result !== "object") return {};
  if (result.task) return { task: normalizeTask(result.task) };
  if (result.message) return { message: result.message };
  if (result.statusUpdate) {
    const s = result.statusUpdate;
    return { statusUpdate: { ...s, status: { ...s.status, state: normalizeState(s.status?.state) } } };
  }
  if (result.artifactUpdate) return { artifactUpdate: result.artifactUpdate };

  switch (result.kind) {
    case "task":
      return { task: normalizeTask(result) };
    case "message":
      return { message: result };
    case "status-update":
      return {
        statusUpdate: { ...result, status: { ...result.status, state: normalizeState(result.status?.state) } },
      };
    case "artifact-update":
      return { artifactUpdate: result };
  }
  // Bare task without `kind` (seen in GetTask-style payloads).
  if (result.id && result.status) return { task: normalizeTask(result) };
  return {};
}

// ---- Helpers ---------------------------------------------------------------

export function partsText(parts: A2aPart[] | undefined): string {
  return (parts ?? [])
    .map((p) => (typeof p.text === "string" ? p.text : ""))
    .filter(Boolean)
    .join("");
}

/** Best-effort answer text from a task (artifacts first, then status message). */
export function taskText(task: A2aTask | undefined): string {
  if (!task) return "";
  const fromArtifacts = (task.artifacts ?? []).map((a) => partsText(a.parts)).filter(Boolean).join("\n\n");
  return fromArtifacts || partsText(task.status?.message?.parts);
}

export function responseText(r: A2aStreamResponse): string {
  if (r.task) return taskText(r.task);
  if (r.message) return partsText(r.message.parts);
  return "";
}

export const TERMINAL_STATES = new Set([
  "TASK_STATE_COMPLETED",
  "TASK_STATE_FAILED",
  "TASK_STATE_CANCELED",
  "TASK_STATE_REJECTED",
]);

/**
 * Folds a stream of A2A events into the current answer text, task id, context id
 * and state. Artifact chunks with `append: true` are concatenated per artifactId.
 */
export class A2aStreamAccumulator {
  taskId?: string;
  contextId?: string;
  state?: string;
  events = 0;
  private artifacts = new Map<string, string>();
  private statusText = "";
  private messageText = "";

  push(evt: A2aStreamResponse): void {
    this.events++;
    if (evt.task) {
      this.taskId = evt.task.id;
      this.contextId = evt.task.contextId;
      this.state = evt.task.status?.state;
      for (const a of evt.task.artifacts ?? []) this.artifacts.set(a.artifactId, partsText(a.parts));
      const s = partsText(evt.task.status?.message?.parts);
      if (s) this.statusText = s;
    } else if (evt.message) {
      this.contextId = evt.message.contextId ?? this.contextId;
      this.messageText += partsText(evt.message.parts);
    } else if (evt.statusUpdate) {
      this.taskId = evt.statusUpdate.taskId;
      this.contextId = evt.statusUpdate.contextId;
      this.state = evt.statusUpdate.status?.state;
      const s = partsText(evt.statusUpdate.status?.message?.parts);
      if (s) this.statusText = s;
    } else if (evt.artifactUpdate) {
      const u = evt.artifactUpdate;
      this.taskId = u.taskId;
      this.contextId = u.contextId;
      const chunk = partsText(u.artifact.parts);
      const prev = this.artifacts.get(u.artifact.artifactId) ?? "";
      this.artifacts.set(u.artifact.artifactId, u.append ? prev + chunk : chunk);
    }
  }

  /** Answer text only (artifacts / agent messages), excluding progress status text. */
  get answer(): string {
    const fromArtifacts = [...this.artifacts.values()].filter(Boolean).join("\n\n");
    return fromArtifacts || this.messageText;
  }

  /** Latest progress text from status updates (e.g. "Looking into it…"). */
  get progress(): string {
    return this.statusText;
  }

  /** Final text: the answer, or the status message if no answer was produced. */
  get text(): string {
    return this.answer || this.statusText;
  }
}
