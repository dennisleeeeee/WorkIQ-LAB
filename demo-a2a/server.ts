import "dotenv/config";
import express from "express";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promises as fs } from "node:fs";
import {
  PublicClientApplication,
  LogLevel,
  type Configuration,
} from "@azure/msal-node";
import {
  WorkIqA2aClient,
  A2aStreamAccumulator,
  type A2aAgentCard,
  type A2aVersion,
} from "../src/a2a.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CACHE_PATH = path.join(__dirname, "..", ".token-cache.json");
const WORK_IQ_SCOPE = "api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask";

const TENANT_ID = process.env.TENANT_ID!;
const CLIENT_ID = process.env.CLIENT_ID!;
const TIME_ZONE = process.env.TIME_ZONE ?? "America/New_York";
const A2A_VERSION = (process.env.A2A_VERSION ?? "1.0") as A2aVersion;
const A2A_AGENT_ID = process.env.A2A_AGENT_ID || undefined;
const PORT = Number(process.env.PORT ?? 3002);

if (!TENANT_ID || !CLIENT_ID) {
  console.error("Missing TENANT_ID / CLIENT_ID in .env");
  process.exit(1);
}

// ---- MSAL public client with file-backed token cache (shared with Demo 1/2) ----
function cachePlugin() {
  return {
    beforeCacheAccess: async (ctx: any) => {
      try {
        ctx.tokenCache.deserialize(await fs.readFile(CACHE_PATH, "utf-8"));
      } catch {
        /* first run */
      }
    },
    afterCacheAccess: async (ctx: any) => {
      if (ctx.cacheHasChanged) {
        await fs.writeFile(CACHE_PATH, ctx.tokenCache.serialize(), "utf-8");
      }
    },
  };
}

const config: Configuration = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
  cache: { cachePlugin: cachePlugin() },
  system: {
    loggerOptions: {
      loggerCallback: (lvl, msg) => {
        if (lvl === LogLevel.Error) console.error(msg);
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Warning,
    },
  },
};
const pca = new PublicClientApplication(config);

async function getSilentToken(): Promise<string | null> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (!accounts.length) return null;
  try {
    const r = await pca.acquireTokenSilent({
      account: accounts[0],
      scopes: [WORK_IQ_SCOPE],
    });
    return r?.accessToken ?? null;
  } catch {
    return null;
  }
}

// ---- Device-code sign-in state (single-user demo) ----
type DeviceFlow = {
  state: "pending" | "done" | "error";
  userCode?: string;
  verificationUri?: string;
  message?: string;
  error?: string;
};
let deviceFlow: DeviceFlow | null = null;

// ---- Work IQ A2A client + current conversation (A2A contextId) ----
const a2a = new WorkIqA2aClient(
  async () => {
    const token = await getSilentToken();
    if (!token) throw new Error("Not signed in.");
    return token;
  },
  { version: A2A_VERSION, agentId: A2A_AGENT_ID }
);
let agentCard: A2aAgentCard | null = null;
let contextId: string | undefined;
let lastTaskId: string | undefined;

// ---------------------------------------------------------------------------
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/status", async (_req, res) => {
  const token = await getSilentToken();
  const accounts = await pca.getTokenCache().getAllAccounts();
  res.json({
    authenticated: !!token,
    account: accounts[0]?.username ?? null,
    endpoint: `A2A ${A2A_VERSION}`,
    timeZone: TIME_ZONE,
  });
});

app.post("/api/login", async (_req, res) => {
  if (deviceFlow?.state === "pending" && deviceFlow.userCode) {
    return res.json(deviceFlow);
  }
  deviceFlow = { state: "pending" };
  await new Promise<void>((resolve) => {
    pca
      .acquireTokenByDeviceCode({
        scopes: [WORK_IQ_SCOPE],
        deviceCodeCallback: (r) => {
          deviceFlow = {
            state: "pending",
            userCode: r.userCode,
            verificationUri: r.verificationUri,
            message: r.message,
          };
          resolve();
        },
      })
      .then(() => {
        deviceFlow = { state: "done" };
      })
      .catch((e) => {
        deviceFlow = { state: "error", error: String(e?.message ?? e) };
        resolve();
      });
  });
  res.json(deviceFlow);
});

app.get("/api/login/poll", (_req, res) => {
  res.json(deviceFlow ?? { state: "error", error: "No login in progress." });
});

app.post("/api/logout", async (_req, res) => {
  const cache = pca.getTokenCache();
  for (const a of await cache.getAllAccounts()) await cache.removeAccount(a);
  await fs.rm(CACHE_PATH, { force: true });
  agentCard = null;
  contextId = lastTaskId = undefined;
  deviceFlow = null;
  res.json({ ok: true });
});

app.post("/api/new-conversation", (_req, res) => {
  contextId = lastTaskId = undefined;
  res.json({ ok: true });
});

/** A2A discovery: GET /.well-known/agent-card.json (cached per sign-in). */
app.get("/api/agent-card", async (_req, res) => {
  try {
    if (!agentCard) agentCard = await a2a.getAgentCard();
    res.json({ card: agentCard, rpcEndpoint: a2a.endpoint, version: a2a.version });
  } catch (e) {
    res.status(500).json({ error: String((e as Error).message) });
  }
});

/** A2A GetTask for the last task in this conversation. */
app.get("/api/task", async (_req, res) => {
  if (!lastTaskId) return res.status(404).json({ error: "No task yet." });
  try {
    res.json(await a2a.getTask(lastTaskId));
  } catch (e) {
    res.status(500).json({ error: String((e as Error).message) });
  }
});

/**
 * A2A SendStreamingMessage, relayed to the browser as NDJSON:
 *   {"type":"event","event":<StreamResponse>}   — one per A2A SSE event
 *   {"type":"done", taskId, contextId, state, text, events}
 *   {"type":"error", error}
 */
app.post("/api/chat", async (req, res) => {
  const text = String(req.body?.text ?? "").trim();
  if (!text) return res.status(400).json({ error: "Empty message." });

  res.setHeader("Content-Type", "application/x-ndjson; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache");
  const write = (obj: unknown) => res.write(JSON.stringify(obj) + "\n");

  try {
    if (!agentCard) agentCard = await a2a.getAgentCard();
    const acc = new A2aStreamAccumulator();
    for await (const evt of a2a.sendStreamingMessage(text, { contextId, timeZone: TIME_ZONE })) {
      acc.push(evt);
      write({ type: "event", event: evt });
    }
    contextId = acc.contextId ?? contextId;
    lastTaskId = acc.taskId ?? lastTaskId;
    write({
      type: "done",
      taskId: acc.taskId,
      contextId,
      state: acc.state,
      text: acc.text,
      events: acc.events,
    });
  } catch (e) {
    write({ type: "error", error: String((e as Error).message) });
  }
  res.end();
});

app.listen(PORT, () => {
  console.log(`\n  Work IQ A2A demo running at  http://localhost:${PORT}\n`);
});
