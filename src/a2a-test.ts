/**
 * Work IQ A2A smoke test + interactive chat.
 *
 *   npm run test:a2a                          # run the protocol smoke test
 *   npm run test:a2a -- --verbose             # also print raw A2A JSON
 *   npm run test:a2a -- --only card,send      # run selected steps
 *   npm run test:a2a -- --prompt "..."        # custom first prompt
 *   npm run a2a                               # interactive streaming chat (REPL)
 *
 * Env: TENANT_ID, CLIENT_ID, TIME_ZONE, plus optional A2A_VERSION (1.0|0.3),
 * A2A_AGENT_ID (target a specific Work IQ agent), WORK_IQ_A2A_ENDPOINT.
 */
import "dotenv/config";
import * as readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { getAccessToken } from "./auth.ts";
import {
  WorkIqA2aClient,
  A2aRpcError,
  A2aStreamAccumulator,
  TERMINAL_STATES,
  responseText,
  taskText,
  type A2aStreamResponse,
  type A2aVersion,
} from "./a2a.ts";

const STEPS = ["card", "send", "multiturn", "gettask", "stream", "cancel"] as const;
type Step = (typeof STEPS)[number];

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value || value.startsWith("your-")) {
    console.error(`Missing env var ${name}. Copy .env.example to .env and fill it in.`);
    process.exit(1);
  }
  return value;
}

function argValue(flag: string): string | undefined {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

const c = {
  green: (s: string) => `\x1b[32m${s}\x1b[0m`,
  red: (s: string) => `\x1b[31m${s}\x1b[0m`,
  yellow: (s: string) => `\x1b[33m${s}\x1b[0m`,
  dim: (s: string) => `\x1b[2m${s}\x1b[0m`,
  bold: (s: string) => `\x1b[1m${s}\x1b[0m`,
};

function truncate(s: string, n = 240): string {
  const one = s.replace(/\s+/g, " ").trim();
  return one.length > n ? one.slice(0, n) + "…" : one;
}

function eventKind(e: A2aStreamResponse): string {
  if (e.task) return `task(${e.task.status?.state})`;
  if (e.message) return "message";
  if (e.statusUpdate) return `statusUpdate(${e.statusUpdate.status?.state})`;
  if (e.artifactUpdate) return `artifactUpdate${e.artifactUpdate.append ? "+append" : ""}`;
  return "unknown";
}

type Outcome = "PASS" | "FAIL" | "WARN" | "SKIP";
interface Result {
  step: string;
  outcome: Outcome;
  detail: string;
  ms: number;
}

class Check extends Error {}
function expect(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Check(msg);
}

async function runSmokeTest(client: WorkIqA2aClient, timeZone: string) {
  const verbose = process.argv.includes("--verbose");
  const only = argValue("--only")?.split(",").map((s) => s.trim()) as Step[] | undefined;
  const prompt = argValue("--prompt") ?? "Hi! In one short sentence, what can you help me with?";
  const results: Result[] = [];
  const ctx: { contextId?: string; taskId?: string } = {};

  const dump = (label: string, obj: unknown) => {
    if (verbose) console.log(c.dim(`  ${label}:\n` + JSON.stringify(obj, null, 2).replace(/^/gm, "    ")));
  };

  async function step(name: Step, title: string, fn: () => Promise<string | { warn: string }>) {
    if (only && !only.includes(name)) {
      results.push({ step: title, outcome: "SKIP", detail: "", ms: 0 });
      return;
    }
    console.log(c.bold(`\n▶ ${title}`));
    const t0 = Date.now();
    try {
      const r = await fn();
      const ms = Date.now() - t0;
      if (typeof r === "string") {
        results.push({ step: title, outcome: "PASS", detail: r, ms });
        console.log(`  ${c.green("PASS")} ${r} ${c.dim(`(${ms} ms)`)}`);
      } else {
        results.push({ step: title, outcome: "WARN", detail: r.warn, ms });
        console.log(`  ${c.yellow("WARN")} ${r.warn} ${c.dim(`(${ms} ms)`)}`);
      }
    } catch (e) {
      const ms = Date.now() - t0;
      const msg = (e as Error).message;
      results.push({ step: title, outcome: "FAIL", detail: msg, ms });
      console.log(`  ${c.red("FAIL")} ${msg} ${c.dim(`(${ms} ms)`)}`);
    }
  }

  console.log(
    c.dim(`A2A-Version ${client.version} · base ${client.baseUrl} · time zone ${timeZone}`)
  );

  // 1) Discovery — the A2A entry point.
  await step("card", "1. Agent discovery (GET /.well-known/agent-card.json)", async () => {
    const card = await client.getAgentCard();
    dump("agentCard", card);
    expect(card?.name, "Agent Card has no name");
    const ifaces = (card.supportedInterfaces ?? []).map((i) => `${i.protocolBinding}@${i.protocolVersion}`);
    return (
      `"${card.name}" v${card.version ?? "?"} · interfaces [${ifaces.join(", ") || "n/a"}] · ` +
      `streaming=${!!card.capabilities?.streaming} · skills=${card.skills?.length ?? 0} · rpc → ${client.endpoint}`
    );
  });

  // 2) Synchronous SendMessage → Task with an Answer artifact.
  await step("send", "2. SendMessage (sync)", async () => {
    console.log(c.dim(`  you > ${prompt}`));
    const r = await client.sendMessage(prompt, { timeZone });
    dump("result", r);
    const text = responseText(r);
    expect(r.task || r.message, "Result contained neither a task nor a message");
    if (r.task) {
      ctx.taskId = r.task.id;
      ctx.contextId = r.task.contextId;
      expect(r.task.id, "Task has no id");
      expect(r.task.contextId, "Task has no contextId");
    } else {
      ctx.contextId = r.message!.contextId;
    }
    expect(text, "No answer text in the response");
    console.log(c.dim(`  agent > ${truncate(text)}`));
    const state = r.task ? r.task.status.state : "direct message";
    return `${state} · taskId=${ctx.taskId ?? "-"} · contextId=${ctx.contextId ?? "-"}`;
  });

  // 3) Multi-turn — reuse contextId; the agent should remember turn 1.
  await step("multiturn", "3. Multi-turn (SendMessage with contextId)", async () => {
    expect(ctx.contextId, "No contextId from step 2 — run the 'send' step first");
    const follow = "Repeat your previous answer in Traditional Chinese, one sentence.";
    console.log(c.dim(`  you > ${follow}`));
    const r = await client.sendMessage(follow, { contextId: ctx.contextId, timeZone });
    dump("result", r);
    const text = responseText(r);
    expect(text, "No answer text in the follow-up response");
    console.log(c.dim(`  agent > ${truncate(text)}`));
    const returnedCtx = r.task?.contextId ?? r.message?.contextId;
    if (returnedCtx && returnedCtx !== ctx.contextId) {
      return { warn: `contextId changed (${ctx.contextId} → ${returnedCtx})` };
    }
    if (r.task?.id) ctx.taskId = r.task.id;
    return `same contextId kept · ${r.task?.status.state ?? "direct message"}`;
  });

  // 4) GetTask — fetch the task back by id.
  await step("gettask", "4. GetTask", async () => {
    expect(ctx.taskId, "No taskId from earlier steps (agent replied with a direct message?)");
    const task = await client.getTask(ctx.taskId);
    dump("task", task);
    expect(task.id === ctx.taskId, `Returned id ${task.id} ≠ requested ${ctx.taskId}`);
    return `id matches · state=${task.status.state} · artifacts=${task.artifacts?.length ?? 0} · "${truncate(taskText(task), 80)}"`;
  });

  // 5) SendStreamingMessage — SSE events folded into the final answer.
  await step("stream", "5. SendStreamingMessage (SSE)", async () => {
    const q = "List three things you can do with my Microsoft 365 data, as short bullets.";
    console.log(c.dim(`  you > ${q}`));
    const acc = new A2aStreamAccumulator();
    const kinds: Record<string, number> = {};
    const t0 = Date.now();
    let firstEventMs = -1;
    for await (const evt of client.sendStreamingMessage(q, { timeZone })) {
      if (firstEventMs < 0) firstEventMs = Date.now() - t0;
      acc.push(evt);
      const k = eventKind(evt);
      kinds[k] = (kinds[k] ?? 0) + 1;
      if (verbose) console.log(c.dim(`    event: ${JSON.stringify(evt)}`));
    }
    expect(acc.events > 0, "Stream produced no events");
    expect(acc.answer, "Stream produced no answer text");
    console.log(c.dim(`  agent > ${truncate(acc.answer)}`));
    const summary = Object.entries(kinds).map(([k, n]) => `${k}×${n}`).join(", ");
    const res = `${acc.events} events [${summary}] · first event ${firstEventMs} ms · final ${acc.state ?? "?"}`;
    if (acc.state && !TERMINAL_STATES.has(acc.state)) return { warn: `${res} (non-terminal final state)` };
    return res;
  });

  // 6) CancelTask on a finished task — spec says TaskNotCancelableError (-32002).
  await step("cancel", "6. CancelTask on a completed task (expects error)", async () => {
    expect(ctx.taskId, "No taskId from earlier steps");
    try {
      const t = await client.cancelTask(ctx.taskId);
      dump("task", t);
      return { warn: `Server accepted cancel; state=${t.status?.state}` };
    } catch (e) {
      if (e instanceof A2aRpcError) {
        if (e.code === -32002) return `TaskNotCancelableError (-32002) as per spec`;
        return { warn: `JSON-RPC error ${e.code}: ${e.message}` };
      }
      throw e;
    }
  });

  // ---- Summary ----
  console.log(c.bold("\n──────── Work IQ A2A test summary ────────"));
  const color = { PASS: c.green, FAIL: c.red, WARN: c.yellow, SKIP: c.dim } as const;
  for (const r of results) {
    console.log(`  ${color[r.outcome](r.outcome.padEnd(4))}  ${r.step}${r.ms ? c.dim(`  ${r.ms} ms`) : ""}`);
  }
  const failed = results.filter((r) => r.outcome === "FAIL").length;
  const passed = results.filter((r) => r.outcome === "PASS").length;
  console.log(`\n  ${passed} passed, ${failed} failed\n`);
  return failed === 0;
}

async function runChat(client: WorkIqA2aClient, timeZone: string) {
  const card = await client.getAgentCard();
  console.log(`Connected to A2A agent "${card.name}" (A2A-Version ${client.version}).`);
  console.log("Streaming chat. Commands: /new (new contextId), /task (GetTask last), /exit.\n");

  const rl = readline.createInterface({ input, output });
  let contextId: string | undefined;
  let taskId: string | undefined;
  try {
    while (true) {
      const prompt = (await rl.question("you > ")).trim();
      if (!prompt) continue;
      if (prompt === "/exit" || prompt === "/quit") break;
      if (prompt === "/new") {
        contextId = taskId = undefined;
        console.log("(new conversation)\n");
        continue;
      }
      if (prompt === "/task") {
        if (!taskId) console.log("(no task yet)\n");
        else console.log(JSON.stringify(await client.getTask(taskId), null, 2) + "\n");
        continue;
      }
      try {
        const acc = new A2aStreamAccumulator();
        let shown = 0;
        process.stdout.write("agent > ");
        for await (const evt of client.sendStreamingMessage(prompt, { contextId, timeZone })) {
          acc.push(evt);
          const text = acc.answer;
          // Print only the newly appended suffix when the text grows monotonically.
          if (text.length > shown && text.startsWith(text.slice(0, shown))) {
            process.stdout.write(text.slice(shown));
            shown = text.length;
          }
        }
        contextId = acc.contextId ?? contextId;
        taskId = acc.taskId ?? taskId;
        console.log(c.dim(`\n  [${acc.state ?? "?"} · ${acc.events} events · ctx ${contextId ?? "-"}]\n`));
      } catch (err) {
        console.error(`\n[error] ${(err as Error).message}\n`);
      }
    }
  } finally {
    rl.close();
  }
}

async function main() {
  const tenantId = requireEnv("TENANT_ID");
  const clientId = requireEnv("CLIENT_ID");
  const timeZone = process.env.TIME_ZONE ?? "America/New_York";
  const version = (argValue("--version") ?? process.env.A2A_VERSION ?? "1.0") as A2aVersion;
  if (version !== "1.0" && version !== "0.3") {
    console.error(`Unsupported A2A version "${version}" (use 1.0 or 0.3).`);
    process.exit(1);
  }

  const client = new WorkIqA2aClient(() => getAccessToken(tenantId, clientId), {
    version,
    agentId: argValue("--agent") ?? process.env.A2A_AGENT_ID ?? undefined,
  });

  if (process.argv.includes("--chat")) {
    await runChat(client, timeZone);
    return;
  }
  const ok = await runSmokeTest(client, timeZone);
  process.exit(ok ? 0 : 1);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
