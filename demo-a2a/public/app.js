const $ = (sel) => document.querySelector(sel);

const appEl = $(".app");
const loginPanel = $("#login-panel");
const chatPanel = $("#chat-panel");
const messagesEl = $("#messages");
const accountName = $("#account-name");
const providerBadge = $("#provider-badge");
const chatInput = $("#chat-input");
const inspector = $("#inspector");
const eventLog = $("#event-log");

let contextId = null;
let taskId = null;

// ---- Boot ----
init();

async function init() {
  const status = await fetch("/api/status").then((r) => r.json());
  if (status.endpoint) {
    providerBadge.textContent = status.endpoint;
    providerBadge.hidden = false;
  }
  if (status.authenticated) showChat(status.account);
  else showLogin();
}

function showLogin() {
  loginPanel.hidden = false;
  loginPanel.style.display = "grid";
  chatPanel.hidden = true;
  chatPanel.style.display = "none";
}

function showChat(account) {
  loginPanel.hidden = true;
  loginPanel.style.display = "none";
  chatPanel.hidden = false;
  chatPanel.style.display = "flex";
  accountName.textContent = account ?? "";
  for (const id of ["#new-btn", "#logout-btn", "#card-btn", "#inspector-btn"]) $(id).hidden = false;
  setInspector(true);
  chatInput.focus();
  discoverAgent();
}

// ---- A2A discovery (Agent Card) ----
let agentCard = null;

async function discoverAgent() {
  logLine("req", "GET /.well-known/agent-card.json", "A2A discovery");
  const res = await fetch("/api/agent-card");
  const data = await res.json();
  if (!res.ok) {
    logLine("err", "Agent Card error", data.error);
    return;
  }
  agentCard = data;
  const c = data.card;
  logLine("task", "AgentCard", `${c.name} v${c.version ?? "?"} · streaming=${!!c.capabilities?.streaming}`, c);
  const line = $("#agent-line");
  line.textContent = `Connected to “${c.name}” · ${data.rpcEndpoint} · A2A ${data.version}`;
  line.hidden = false;
}

$("#card-btn").addEventListener("click", async () => {
  if (!agentCard) await discoverAgent();
  $("#card-json").textContent = JSON.stringify(agentCard?.card ?? {}, null, 2);
  $("#card-dialog").showModal();
});
$("#card-close").addEventListener("click", () => $("#card-dialog").close());

// ---- Sign in (device code) ----
$("#login-btn").addEventListener("click", startLogin);

async function startLogin() {
  const btn = $("#login-btn");
  const errEl = $("#login-error");
  errEl.hidden = true;
  btn.disabled = true;
  btn.textContent = "Getting sign-in code…";

  const flow = await fetch("/api/login", { method: "POST" }).then((r) => r.json());
  if (flow.state === "error") {
    errEl.textContent = flow.error;
    errEl.hidden = false;
    btn.disabled = false;
    btn.textContent = "Sign in";
    return;
  }

  $("#device-code").hidden = false;
  $("#device-uri").textContent = flow.verificationUri;
  $("#device-uri").href = flow.verificationUri;
  $("#device-user-code").textContent = flow.userCode;
  btn.textContent = "Waiting for sign-in…";
  pollLogin();
}

$("#copy-code").addEventListener("click", () => {
  navigator.clipboard.writeText($("#device-user-code").textContent);
  $("#copy-code").textContent = "Copied";
  setTimeout(() => ($("#copy-code").textContent = "Copy"), 1500);
});

async function pollLogin() {
  const errEl = $("#login-error");
  const timer = setInterval(async () => {
    const flow = await fetch("/api/login/poll").then((r) => r.json());
    if (flow.state === "done") {
      clearInterval(timer);
      const status = await fetch("/api/status").then((r) => r.json());
      showChat(status.account);
    } else if (flow.state === "error") {
      clearInterval(timer);
      errEl.textContent = flow.error;
      errEl.hidden = false;
      $("#login-btn").disabled = false;
      $("#login-btn").textContent = "Retry";
      $("#device-code").hidden = true;
    }
  }, 2500);
}

// ---- Chat ----
$("#chat-form").addEventListener("submit", (e) => {
  e.preventDefault();
  const text = chatInput.value.trim();
  if (text) {
    chatInput.value = "";
    updateSendState();
    sendMessage(text);
  }
});

chatInput.addEventListener("input", updateSendState);
function updateSendState() {
  chatPanel.classList.toggle("can-send", chatInput.value.trim().length > 0);
}

$("#suggestions").addEventListener("click", (e) => {
  const btn = e.target.closest(".suggestion");
  if (btn) sendMessage(btn.querySelector("span").textContent);
});

/** Folds A2A StreamResponse events into answer text (mirrors src/a2a.ts). */
function createAccumulator() {
  const artifacts = new Map();
  const partsText = (parts) => (parts ?? []).map((p) => (typeof p.text === "string" ? p.text : "")).join("");
  return {
    progress: "",
    state: null,
    push(evt) {
      if (evt.task) {
        this.state = evt.task.status?.state;
        for (const a of evt.task.artifacts ?? []) artifacts.set(a.artifactId, partsText(a.parts));
        const s = partsText(evt.task.status?.message?.parts);
        if (s) this.progress = s;
      } else if (evt.statusUpdate) {
        this.state = evt.statusUpdate.status?.state;
        const s = partsText(evt.statusUpdate.status?.message?.parts);
        if (s) this.progress = s;
      } else if (evt.artifactUpdate) {
        const u = evt.artifactUpdate;
        const prev = artifacts.get(u.artifact.artifactId) ?? "";
        const chunk = partsText(u.artifact.parts);
        artifacts.set(u.artifact.artifactId, u.append ? prev + chunk : chunk);
      } else if (evt.message) {
        const id = evt.message.messageId ?? "message";
        artifacts.set(id, (artifacts.get(id) ?? "") + partsText(evt.message.parts));
      }
    },
    get answer() {
      return [...artifacts.values()].filter(Boolean).join("\n\n");
    },
  };
}

async function sendMessage(text) {
  chatPanel.classList.add("has-messages");
  addMessage(text, "user");
  const bot = addMessage("Sending A2A task…", "bot", true);
  setSending(true);
  logLine("req", "→ SendStreamingMessage", `“${text}”` + (contextId ? ` · contextId ${short(contextId)}` : " · new context"));

  const acc = createAccumulator();
  let rendered = "";
  let raf = 0;
  const render = () => {
    raf = 0;
    const answer = acc.answer;
    if (answer && answer !== rendered) {
      bot.classList.remove("thinking");
      bot.classList.add("streaming");
      bot.innerHTML = marked.parse(answer);
      rendered = answer;
    } else if (!answer && acc.progress) {
      bot.innerHTML = `<span class="progress"></span>`;
      bot.firstChild.textContent = acc.progress;
    }
    messagesEl.scrollTop = messagesEl.scrollHeight;
  };

  try {
    const res = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    });
    if (!res.ok || !res.body) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || `HTTP ${res.status}`);
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let done = null;
    while (true) {
      const { value, done: eof } = await reader.read();
      if (eof) break;
      buffer += decoder.decode(value, { stream: true });
      let idx;
      while ((idx = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, idx).trim();
        buffer = buffer.slice(idx + 1);
        if (!line) continue;
        const msg = JSON.parse(line);
        if (msg.type === "event") {
          acc.push(msg.event);
          logEvent(msg.event);
          if (!raf) raf = requestAnimationFrame(render);
        } else if (msg.type === "done") {
          done = msg;
        } else if (msg.type === "error") {
          throw new Error(msg.error);
        }
      }
    }

    if (raf) cancelAnimationFrame(raf);
    render();
    bot.classList.remove("streaming", "thinking");
    if (!acc.answer) bot.innerHTML = marked.parse(done?.text || "(no answer)");
    bot.querySelectorAll("a").forEach((a) => (a.target = "_blank"));

    if (done) {
      contextId = done.contextId ?? contextId;
      taskId = done.taskId ?? taskId;
      updateCtxLine();
      const meta = document.createElement("div");
      meta.className = "msg-meta";
      meta.textContent = `${done.state ?? "?"} · ${done.events} A2A events · task ${short(done.taskId)}`;
      bot.appendChild(meta);
      logLine(done.state === "TASK_STATE_COMPLETED" ? "task" : "err", "■ stream closed", `${done.state} · ${done.events} events`);
    }
  } catch (err) {
    bot.classList.remove("streaming", "thinking");
    bot.textContent = "⚠️ " + err.message;
    logLine("err", "error", err.message);
  } finally {
    setSending(false);
    chatInput.focus();
  }
}

function addMessage(text, role, thinking = false) {
  const el = document.createElement("div");
  el.className = `msg ${role}` + (thinking ? " thinking" : "");
  el.textContent = text;
  messagesEl.appendChild(el);
  messagesEl.scrollTop = messagesEl.scrollHeight;
  return el;
}

function setSending(on) {
  $("#send-btn").disabled = on;
  chatInput.disabled = on;
}

// ---- Inspector ----
function short(id) {
  return id ? String(id).slice(0, 8) + "…" : "—";
}

function updateCtxLine() {
  $("#ctx-line").textContent = `contextId: ${contextId ?? "—"} · taskId: ${taskId ?? "—"}`;
}

function setInspector(on) {
  inspector.hidden = !on;
  appEl.classList.toggle("with-inspector", on);
}
$("#inspector-btn").addEventListener("click", () => setInspector(inspector.hidden));
$("#clear-log").addEventListener("click", () => (eventLog.innerHTML = ""));

$("#task-btn").addEventListener("click", async () => {
  logLine("req", "→ GetTask", taskId ? short(taskId) : "(no task yet)");
  const res = await fetch("/api/task");
  const data = await res.json();
  if (!res.ok) logLine("err", "GetTask error", data.error);
  else logLine("task", "Task", `${data.status?.state} · artifacts=${data.artifacts?.length ?? 0}`, data);
});

function logLine(kind, label, summary, raw) {
  const li = document.createElement("li");
  const head = `<span class="k ${kind}"></span> <span class="s"></span>`;
  if (raw !== undefined) {
    li.innerHTML = `<details><summary>${head}</summary><pre></pre></details>`;
    li.querySelector("pre").textContent = JSON.stringify(raw, null, 2);
  } else {
    li.innerHTML = head;
  }
  li.querySelector(".k").textContent = label;
  li.querySelector(".s").textContent = summary ?? "";
  eventLog.appendChild(li);
  eventLog.scrollTop = eventLog.scrollHeight;
  return li;
}

/** Log one A2A StreamResponse; consecutive append chunks collapse into one row. */
let chunkRow = null;
function logEvent(evt) {
  const kind = Object.keys(evt)[0];
  if (kind === "artifactUpdate" && evt.artifactUpdate.append && chunkRow?.artifactId === evt.artifactUpdate.artifact.artifactId) {
    chunkRow.count++;
    chunkRow.li.querySelector(".k").textContent = `artifactUpdate ×${chunkRow.count}`;
    chunkRow.li.querySelector("pre").textContent = JSON.stringify(evt, null, 2);
    return;
  }
  chunkRow = null;

  let summary = "";
  if (kind === "task") summary = `${evt.task.status?.state} · id ${short(evt.task.id)} · ctx ${short(evt.task.contextId)}`;
  else if (kind === "statusUpdate") {
    const p = (evt.statusUpdate.status?.message?.parts ?? []).map((x) => x.text).filter(Boolean).join("");
    summary = `${evt.statusUpdate.status?.state}${p ? ` · “${p}”` : ""}`;
  } else if (kind === "artifactUpdate") summary = `${evt.artifactUpdate.artifact.name ?? "artifact"}${evt.artifactUpdate.append ? " (append)" : ""}`;
  else if (kind === "message") summary = "agent message";

  const li = logLine(kind, kind, summary, evt);
  if (kind === "artifactUpdate" && evt.artifactUpdate.append) {
    chunkRow = { li, count: 1, artifactId: evt.artifactUpdate.artifact.artifactId };
  }
}

// ---- Header actions ----
$("#new-btn").addEventListener("click", async () => {
  await fetch("/api/new-conversation", { method: "POST" });
  contextId = taskId = null;
  updateCtxLine();
  logLine("req", "— new conversation —", "next message starts a new contextId");
  messagesEl.innerHTML = "";
  chatPanel.classList.remove("has-messages", "can-send");
  chatInput.value = "";
  chatInput.focus();
});

$("#logout-btn").addEventListener("click", async () => {
  await fetch("/api/logout", { method: "POST" });
  location.reload();
});
