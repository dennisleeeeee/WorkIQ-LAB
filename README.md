# WorkIQ LAB

Three browser demos for the **Microsoft 365 Work IQ API**, showing the three ways an
app can use Work IQ:

| Demo | Folder | What it is | Work IQ surface |
| --- | --- | --- | --- |
| **Demo 1 — Chat** | [`demo-rest/`](./demo-rest) | A website that calls the Work IQ **REST Chat API** directly. Pure grounded Q&A. | `POST /conversations/{id}/chat` |
| **Demo 2 — Agent** | [`demo-agent/`](./demo-agent) | A website with a custom **LangChain.js agent** whose tools are the Work IQ **MCP** server. The LLM decides when to answer vs. take an action. | MCP `ask` + `do_action` |
| **Demo 3 — A2A** | [`demo-a2a/`](./demo-a2a) | A website that talks to Work IQ as a peer agent over the **A2A (Agent-to-Agent) protocol**: Agent Card discovery + streaming JSON-RPC tasks, with a live protocol inspector. | A2A `SendStreamingMessage`, `GetTask` |

All demos sign the user in with Microsoft Entra (delegated **device-code flow**) and keep
tokens server-side. Work IQ only supports **delegated** permissions (signed-in user).

```
User ──▶ Demo 1 website ──▶ Work IQ REST Chat API      (ask questions, text back)

User ──▶ Demo 2 website ──▶ LangChain agent (LLM brain) ──▶ Work IQ MCP
                                                              ├─ ask        → grounded chat
                                                              └─ do_action  → real actions (send mail)

User ──▶ Demo 3 website ──▶ Work IQ A2A agent  (JSON-RPC 2.0 + SSE)
                              ├─ GET  /.well-known/agent-card.json   → discovery
                              └─ POST SendStreamingMessage / GetTask → tasks, contextId multi-turn
```

---

## Prerequisites

1. **Work IQ enabled** in your tenant by an admin —
   [Enable Work IQ](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/enable-work-iq).
2. A **Microsoft Entra app registration** (public client):
   - Delegated API permission **`WorkIQAgent.Ask`** (Work IQ API, app ID URI
     `api://workiq.svc.cloud.microsoft`), with **admin consent**.
   - **Authentication → Allow public client flows → Yes** (for device-code flow).
3. **Node.js 18+**.
4. **Demo 2 only:** an LLM with tool-calling (Microsoft Foundry / Azure OpenAI /
   OpenAI / GitHub Models).

---

## Setup

```bash
npm install
cp .env.example .env      # then edit .env (see below)
```

### `.env` — shared (all demos)

| Variable | Description |
| --- | --- |
| `TENANT_ID` | Directory (tenant) ID, or `organizations` |
| `CLIENT_ID` | Application (client) ID of your Entra app registration |
| `USE_BETA` | `true` to hit the `/beta` Chat endpoint, else `false` (Demo 1) |
| `TIME_ZONE` | IANA time zone sent as `locationHint` (e.g. `America/New_York`) |
| `A2A_VERSION` | *(optional, Demo 3)* A2A protocol version header, `1.0` (default) or `0.3` |
| `A2A_AGENT_ID` | *(optional, Demo 3)* target a specific Work IQ agent instead of the default Copilot agent |
| `WORK_IQ_A2A_ENDPOINT` | *(optional, Demo 3)* defaults to `https://workiq.svc.cloud.microsoft/a2a` |

### `.env` — Demo 2 agent brain

Pick a provider with `AGENT_PROVIDER` and fill its variables:

```bash
# Microsoft Foundry (OpenAI-compatible /openai/v1 endpoint + API key)
AGENT_PROVIDER=foundry
FOUNDRY_ENDPOINT=https://<resource>.services.ai.azure.com/openai/v1   # note: /openai/v1, NOT /responses
FOUNDRY_DEPLOYMENT=<deployment-name>                                  # e.g. gpt-4o / gpt-chat-latest
FOUNDRY_API_KEY=<key from the Foundry "Call model" panel>

# — or — OpenAI          AGENT_PROVIDER=openai   OPENAI_API_KEY=...  AGENT_MODEL=gpt-4o-mini
# — or — GitHub Models   AGENT_PROVIDER=github   GITHUB_TOKEN=...    AGENT_MODEL=gpt-4o-mini
# — or — Azure OpenAI    AGENT_PROVIDER=azure    AZURE_OPENAI_API_KEY / _INSTANCE_NAME / _DEPLOYMENT_NAME / _VERSION
```

> Tip: prefer a fast chat model (e.g. `gpt-4o`). Reasoning models like `gpt-5` are
> slow (30–40 s/turn) and more prone to rate limits.

---

## How to start the demos

### Demo 1 — Work IQ REST Chat  →  http://localhost:3000

```bash
npm run web:rest
```

### Demo 2 — Agent (LangChain + Work IQ MCP)  →  http://localhost:3001

```bash
npm run web:agent
```

### Demo 3 — A2A (Agent-to-Agent protocol)  →  http://localhost:3002

```bash
npm run web:a2a
```

The right-hand **Inspector** shows every A2A event (`task`, `statusUpdate`,
`artifactUpdate`…), the current `contextId` / `taskId`, and a **GetTask** button.
**Agent Card** shows the discovered card JSON.

Run them at once in separate terminals. Change the port with `PORT`, e.g.
`PORT=8080 npm run web:rest`.

**First-time sign-in:** the page shows a device code and a link. Open
[https://microsoft.com/devicelogin](https://microsoft.com/devicelogin), enter the
code, and sign in. The token is cached in `.token-cache.json` (shared by all
demos), so later runs go straight to the chat screen.

### CLI (optional, Demo 1 style)

```bash
npm start            # interactive REST chat REPL
npm start -- --stream
```

REPL commands: `/stream`, `/sync`, `/exit`.

### A2A test (CLI, Demo 3 style)

```bash
npm run test:a2a                     # smoke test: card, send, multiturn, gettask, stream, cancel
npm run test:a2a -- --verbose        # also print raw JSON-RPC results / events
npm run test:a2a -- --only card,stream --prompt "What meetings do I have today?"
npm run test:a2a -- --version 0.3    # test the legacy v0.3 method names
npm run a2a                          # interactive streaming A2A chat REPL (/new, /task, /exit)
```

Each step prints PASS / FAIL / WARN / SKIP; the process exits `1` if any step fails.

### A2A — the minimum an engineer needs

[`src/a2a-minimal.ts`](./src/a2a-minimal.ts) (~50 lines, plain `fetch`) is the
smallest working Work IQ A2A call: `npm run a2a:min -- "your question"`.

1. **Entra app** (public client) with delegated `WorkIQAgent.Ask` + admin consent — same as REST/MCP.
2. **Delegated user token** for `api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask` (no app-only).
3. *(optional)* `GET https://workiq.svc.cloud.microsoft/a2a/.well-known/agent-card.json` — discovery.
4. `POST https://workiq.svc.cloud.microsoft/a2a/` with headers `Authorization: Bearer …`,
   `A2A-Version: 1.0` and JSON-RPC body `{"method":"SendMessage","params":{"message":{"role":"ROLE_USER","messageId":"<uuid>","parts":[{"text":"…"}]}}}`.
5. Read the answer from `result.task.artifacts[].parts[].text`; pass `contextId` back for multi-turn.

---

## Using the demos

**Demo 1** — ask questions grounded in your Microsoft 365 data, e.g.
"我今天有哪些會議？", "摘要我未讀的 Outlook 郵件". It only returns text.

**Demo 2** — same questions work (the agent calls the MCP `ask` tool), **and** it can
take actions from natural language, e.g. "發一封提醒信件給我自己記得報帳" → the agent
calls `do_action /me/sendMail`.

### Actions need a tenant policy (Demo 2)

Work IQ MCP **blocks all mutations (create/update/delete/action, incl. sending mail)
by default**. To allow `do_action /me/sendMail`, a tenant admin must enable the mail
mutation scenario in the Microsoft 365 admin center:
**Agents → Tools → Work IQ MCP → Policy** (can take up to 24 h to apply). Until then,
send attempts return `Path is not in the policy allowlist`. No code change is needed
once the policy is enabled. See
[Policy governance for Work IQ MCP](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/policy-governance-mcp).

---

## Manual REST / A2A testing (no Node)

Open [`requests.http`](./requests.http) with the VS Code **REST Client** extension,
paste a delegated access token, and fire requests directly at the Work IQ REST API
or the A2A endpoint (Agent Card, `SendMessage`, `SendStreamingMessage`, `GetTask`,
`CancelTask`).

---

## Project layout

```
src/                     # shared clients
  auth.ts                #   MSAL device-code flow + token cache (WorkIQAgent.Ask)
  workiq.ts              #   Work IQ REST Chat client (Demo 1)
  mcp.ts                 #   Work IQ MCP client (Demo 2)
  a2a.ts                 #   Work IQ A2A client: Agent Card, JSON-RPC, SSE streaming (Demo 3)
  a2a-test.ts            #   A2A smoke test + streaming chat REPL
  a2a-minimal.ts         #   smallest possible A2A call (plain fetch)
  index.ts               #   CLI (Demo 1 chat REPL)
demo-rest/               # Demo 1 website — HTML + Work IQ REST Chat API
  server.ts              #   Express backend (auth + Chat proxy)
  public/                #   index.html, app.js, styles.css
demo-agent/              # Demo 2 website — HTML + LangChain agent + Work IQ MCP
  server.ts              #   Express backend (auth + agent)
  agent.ts               #   LangChain tool-calling agent over MCP (ask / do_action)
  public/                #   index.html, app.js, styles.css
demo-a2a/                # Demo 3 website — HTML + Work IQ A2A protocol
  server.ts              #   Express backend (auth + A2A stream relay as NDJSON)
  public/                #   index.html, app.js (chat + protocol inspector), styles.css
scripts/
  provisionWorkIqSp.ts   #   one-off: create the Work IQ service principal (Node)
  provision-workiq-sp.ps1#   one-off: same, PowerShell
requests.http            # manual REST + A2A request collection (Demo 1 & 3)
```

### npm scripts

| Script | Does |
| --- | --- |
| `npm run web:rest` | Start Demo 1 (REST Chat) — default port 3000 |
| `npm run web:agent` | Start Demo 2 (Agent) — default port 3001 |
| `npm run web:a2a` | Start Demo 3 (A2A) — default port 3002 |
| `npm start` | CLI Demo 1 chat REPL |
| `npm run test:a2a` | A2A protocol smoke test (6 steps) |
| `npm run a2a` | CLI A2A streaming chat REPL |
| `npm run a2a:min` | Minimal single-file A2A call (`src/a2a-minimal.ts`) |
| `npm run typecheck` | `tsc --noEmit` |

---

## Work IQ API cheat-sheet

| Surface | Endpoint / call | Used by |
| --- | --- | --- |
| REST Chat | `POST /conversations`, `/conversations/{id}/chat`, `/chatOverStream` | Demo 1 |
| MCP | `https://workiq.svc.cloud.microsoft/mcp` — tools `ask`, `do_action`, `fetch`, … | Demo 2 |
| A2A | `https://workiq.svc.cloud.microsoft/a2a` — `GET /.well-known/agent-card.json`, JSON-RPC `SendMessage`, `SendStreamingMessage`, `GetTask`, `CancelTask` (header `A2A-Version: 1.0`) | Demo 3 |
| Auth scope | `api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask` (delegated) | all |

The Work IQ **REST API has no action/tool endpoints** — it's chat only. Actions
(the "Work IQ Tool API") are exposed **only through MCP** today.

## Docs

- [Work IQ API overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/api-overview)
- [Work IQ REST Chat API](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/rest/overview)
- [Work IQ MCP overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/overview) ·
  [tool reference](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/tool-reference) ·
  [policy governance](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/policy-governance-mcp)
- [Work IQ A2A overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/a2a/overview) ·
  [A2A protocol spec](https://a2a-protocol.org/latest/specification/)
- [Enable Work IQ](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/enable-work-iq)
