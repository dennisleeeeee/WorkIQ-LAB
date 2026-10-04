/**
 * Demo 2 — a custom LangChain.js agent whose tools are backed by the Work IQ
 * MCP server. The LLM "brain" decides, per turn, whether to:
 *   - call `workiq_ask`       → Work IQ MCP `ask` tool (grounded Copilot chat)
 *   - call `workiq_send_mail`  → Work IQ MCP `do_action /me/sendMail` (a real action)
 *   - answer directly
 *
 * This is the tool-calling agent pattern: bind tools to the model, loop while
 * the model emits tool calls, feed results back as ToolMessages.
 */

import { z } from "zod";
import { tool } from "@langchain/core/tools";
import {
  AIMessage,
  HumanMessage,
  SystemMessage,
  ToolMessage,
  type BaseMessage,
} from "@langchain/core/messages";
import { AzureChatOpenAI, ChatOpenAI } from "@langchain/openai";
import { WorkIqMcpClient, type McpToolResult } from "../src/mcp.ts";

/** Build the chat model "brain" from env. Provider is selectable. */
export function buildModel() {
  const provider = (process.env.AGENT_PROVIDER ?? "azure").toLowerCase();

  if (provider === "foundry") {
    // Microsoft Foundry model via its OpenAI-compatible /openai/v1 endpoint,
    // authenticated with an API key.
    const endpoint = process.env.FOUNDRY_ENDPOINT;
    if (!endpoint) {
      throw new Error("FOUNDRY_ENDPOINT is required when AGENT_PROVIDER=foundry.");
    }
    return new ChatOpenAI({
      model: process.env.FOUNDRY_DEPLOYMENT ?? "gpt-5",
      apiKey: process.env.FOUNDRY_API_KEY,
      configuration: { baseURL: endpoint },
      // Fail fast instead of hanging on the SDK's exponential backoff when the
      // deployment is rate limited (429).
      maxRetries: 1,
      timeout: 60000,
    });
  }

  if (provider === "azure") {
    return new AzureChatOpenAI({
      azureOpenAIApiKey: process.env.AZURE_OPENAI_API_KEY,
      azureOpenAIApiInstanceName: process.env.AZURE_OPENAI_API_INSTANCE_NAME,
      azureOpenAIApiDeploymentName: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME,
      azureOpenAIApiVersion:
        process.env.AZURE_OPENAI_API_VERSION ?? "2024-10-21",
      temperature: 0,
    });
  }

  if (provider === "github") {
    return new ChatOpenAI({
      model: process.env.AGENT_MODEL ?? "gpt-4o-mini",
      apiKey: process.env.GITHUB_TOKEN,
      configuration: { baseURL: "https://models.inference.ai.azure.com" },
      temperature: 0,
    });
  }

  // openai
  return new ChatOpenAI({
    model: process.env.AGENT_MODEL ?? "gpt-4o-mini",
    apiKey: process.env.OPENAI_API_KEY,
    temperature: 0,
  });
}

/** Convert an MCP tool result into text the model can read. */
function resultToText(result: McpToolResult): string {
  const text = (result.content ?? [])
    .filter((c) => c.type === "text" && c.text)
    .map((c) => c.text)
    .join("\n");
  if (text) return text;
  if (result.structuredContent) return JSON.stringify(result.structuredContent);
  return result.isError ? "Tool returned an error." : "(no content)";
}

export class WorkIqAgent {
  private readonly model: ReturnType<typeof buildModel>;
  private boundModel: any;
  private tools: ReturnType<typeof tool>[] = [];
  private history: BaseMessage[] = [];
  private queue: Promise<unknown> = Promise.resolve();

  private constructor(
    private readonly mcp: WorkIqMcpClient,
    private readonly me: string | null,
    private readonly timeZone: string
  ) {
    this.model = buildModel();
    this.reset();
  }

  /**
   * Create an agent whose tools are ALL the tools the Work IQ MCP server exposes
   * (ask, list_agents, fetch, create/update/delete_entity, do_action,
   * call_function, get_schema, search_paths). Tools are generated dynamically
   * from the MCP tool list (name + description + JSON-Schema), so the agent
   * always matches whatever the server offers.
   */
  static async create(
    mcp: WorkIqMcpClient,
    me: string | null,
    timeZone: string
  ): Promise<WorkIqAgent> {
    const agent = new WorkIqAgent(mcp, me, timeZone);
    await agent.initTools();
    return agent;
  }

  private async initTools(): Promise<void> {
    const { tools } = await this.mcp.listTools();
    console.log(`[mcp] tools (${tools.length}): ${tools.map((t) => t.name).join(", ")}`);
    this.tools = tools.map((t) =>
      tool(
        async (args: Record<string, unknown>) => {
          // Inject the user's time zone into `ask` when the model omits it.
          const finalArgs =
            t.name === "ask" && args && !("timeZone" in args)
              ? { ...args, timeZone: this.timeZone }
              : args ?? {};
          const result = await this.mcp.callTool(t.name, finalArgs);
          return resultToText(result);
        },
        {
          name: t.name,
          description: t.description ?? t.name,
          schema: (t.inputSchema as any) ?? z.object({}),
        }
      )
    );
    this.boundModel = this.model.bindTools(this.tools);
  }

  /** Reset the conversation history (keeps the system prompt). */
  reset(): void {
    this.history = [
      new SystemMessage(
        [
          "You are a helpful Microsoft 365 assistant for the signed-in user.",
          this.me
            ? `The user's own email address is ${this.me}. When they say "myself" / "me" / "我自己", use this address.`
            : "",
          "You have Work IQ tools that operate on Microsoft 365 resource paths:",
          "- ask: natural-language questions grounded in the user's data (emails, meetings, files, chats, people). Prefer this for information.",
          "- fetch / call_function: read entities or compute results by resource path.",
          "- create_entity / update_entity / delete_entity / do_action: create, change, delete, or perform actions (e.g. do_action /me/sendMail to send email; /me/chats/{id}/messages for Teams; create_entity /me/events for a meeting).",
          "- search_paths and get_schema: discover valid paths and their request schema. When unsure of the exact path or body for a create/update/do_action, call search_paths and/or get_schema FIRST, then call the action tool.",
          "Base answers on tool results; if a read returns nothing, say there are none — never claim you cannot access Microsoft 365.",
          "Some mutating actions may be blocked by tenant policy. If a tool returns an access/policy error (e.g. 'not in the policy allowlist'), tell the user the action is blocked by tenant policy.",
          "After an action, briefly confirm what you did.",
        ]
          .filter(Boolean)
          .join(" ")
      ),
    ];
  }

  /**
   * Ensure every assistant tool_call id is immediately followed by ToolMessages
   * answering it. Repairs any dangling tool_call left over from a prior turn so
   * the model API never rejects the history with a 400.
   */
  private repairHistory(): void {
    const repaired: BaseMessage[] = [];
    for (let i = 0; i < this.history.length; i++) {
      const m: any = this.history[i];
      repaired.push(m);

      const ids = new Set<string>(
        [
          ...(m.tool_calls ?? []).map((c: any) => c.id),
          ...(m.additional_kwargs?.tool_calls ?? []).map((c: any) => c.id),
        ].filter(Boolean)
      );
      if (!ids.size) continue;

      // Which of these ids are answered by the contiguous following ToolMessages?
      const answered = new Set<string>();
      for (let j = i + 1; j < this.history.length; j++) {
        const next: any = this.history[j];
        if (next?.constructor?.name !== "ToolMessage") break;
        answered.add(next.tool_call_id);
      }
      for (const id of ids) {
        if (!answered.has(id)) {
          repaired.push(
            new ToolMessage({ content: "(no result)", tool_call_id: id })
          );
        }
      }
    }
    this.history = repaired;
  }

  /** Run one user turn through the tool-calling loop and return the reply text. */
  async chat(userText: string): Promise<string> {
    // Serialize turns: concurrent requests must not interleave writes to the
    // shared history, or the tool-call/tool-result pairing gets corrupted.
    const run = this.queue.then(() => this.runTurn(userText));
    this.queue = run.then(
      () => undefined,
      () => undefined
    );
    return run;
  }

  private async runTurn(userText: string): Promise<string> {
    this.history.push(new HumanMessage(userText));
    this.repairHistory();
    const toolMap = new Map(this.tools.map((t) => [t.name, t]));

    for (let step = 0; step < 8; step++) {
      let ai: AIMessage;
      try {
        ai = await this.boundModel.invoke(this.history);
      } catch (err: any) {
        const status = err?.status ?? err?.response?.status;
        if (status === 429) {
          return "⚠️ 模型目前被限流(429 rate limit)。這通常是 Foundry 部署的 TPM/RPM 額度不足或短時間請求過多。請稍候再試,或在 Foundry 提高該部署的配額。";
        }
        if (err?.name === "TimeoutError" || /timeout/i.test(String(err?.message))) {
          return "⚠️ 模型回應逾時。請再試一次,或改用回應較快的部署。";
        }
        throw err;
      }
      this.history.push(ai);

      const calls = ai.tool_calls ?? [];
      // Tool calls whose arguments failed to parse land here; they still carry an
      // id that needs a response.
      const invalid = (ai as any).invalid_tool_calls ?? [];
      // The raw tool_calls actually sent back to the API. Every one of these ids
      // MUST be answered by a ToolMessage or the next request 400s.
      const rawIds: string[] = (ai.additional_kwargs?.tool_calls ?? [])
        .map((c: any) => c.id)
        .filter(Boolean);

      if (!calls.length && !invalid.length && !rawIds.length) {
        return typeof ai.content === "string"
          ? ai.content
          : JSON.stringify(ai.content);
      }

      const answered = new Set<string>();

      for (const call of calls) {
        console.log(
          `[tool] step=${step} → ${call.name} ${JSON.stringify(call.args).slice(0, 500)}`
        );
        const selected = toolMap.get(call.name);
        let output: string;
        try {
          output = selected
            ? String(await selected.invoke(call.args))
            : `Unknown tool: ${call.name}`;
        } catch (err) {
          output = `Tool error: ${(err as Error).message}`;
        }
        console.log(`[tool] step=${step} ← ${call.name}: ${output.slice(0, 500)}`);
        this.history.push(
          new ToolMessage({ content: output, tool_call_id: call.id! })
        );
        if (call.id) answered.add(call.id);
      }

      for (const bad of invalid) {
        this.history.push(
          new ToolMessage({
            content: `Invalid tool call for "${bad.name ?? "unknown"}": ${
              bad.error ?? "arguments could not be parsed"
            }. Please retry with valid JSON arguments.`,
            tool_call_id: bad.id,
          })
        );
        if (bad.id) answered.add(bad.id);
      }

      // Safety net: guarantee every raw tool_call id got a ToolMessage.
      for (const id of rawIds) {
        if (!answered.has(id)) {
          this.history.push(
            new ToolMessage({
              content: "Tool call could not be processed.",
              tool_call_id: id,
            })
          );
        }
      }
    }

    return "(agent stopped after too many steps)";
  }
}
