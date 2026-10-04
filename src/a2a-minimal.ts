/**
 * Minimal Work IQ A2A call — the bare minimum, no SDK, no framework.
 *
 *   npm run a2a:min -- "What meetings do I have today?"
 *
 * What it takes:
 *   1. A delegated token for api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask
 *   2. GET  {A2A}/.well-known/agent-card.json      (discovery — optional but standard)
 *   3. POST {A2A}/  JSON-RPC "SendMessage"          (header A2A-Version: 1.0)
 *   4. Read the answer from result.task.artifacts[].parts[].text
 */
import "dotenv/config";
import { randomUUID } from "node:crypto";
import { getAccessToken } from "./auth.ts";

const A2A = "https://workiq.svc.cloud.microsoft/a2a";
const question = process.argv.slice(2).join(" ") || "Hi! In one sentence, what can you help me with?";

// 1. Token (cached after first device-code sign-in)
const token = await getAccessToken(process.env.TENANT_ID!, process.env.CLIENT_ID!);
const headers = {
  Authorization: `Bearer ${token}`,
  "A2A-Version": "1.0",
  "Content-Type": "application/json",
};

// 2. Discover the agent
const card = await fetch(`${A2A}/.well-known/agent-card.json`, { headers }).then((r) => r.json() as any);
console.log(`Agent: ${card.name} — ${card.description}\n`);

// 3. Send one message (JSON-RPC 2.0)
const res = await fetch(`${A2A}/`, {
  method: "POST",
  headers,
  body: JSON.stringify({
    jsonrpc: "2.0",
    id: 1,
    method: "SendMessage",
    params: {
      message: { role: "ROLE_USER", messageId: randomUUID(), parts: [{ text: question }] },
    },
  }),
}).then((r) => r.json() as any);

if (res.error) throw new Error(`A2A error ${res.error.code}: ${res.error.message}`);

// 4. Answer text lives in the task's artifacts
const task = res.result.task;
const answer = task.artifacts
  ?.flatMap((a: any) => a.parts)
  .map((p: any) => p.text ?? "")
  .join("");
console.log(`you   > ${question}\nagent > ${answer}\n`);
console.log(`state=${task.status.state} taskId=${task.id} contextId=${task.contextId}`);
