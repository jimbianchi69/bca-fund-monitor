// bca-fund-monitor — Phase 1: Broadridge Fiduciary Focus Toolkit webhook receiver
//
// Purpose of this phase is discovery: accept every webhook the Toolkit sends,
// persist the complete payload (headers, query, body) untouched, and expose a
// token-protected review endpoint so we can learn the real event schema before
// building the quarterly monitoring run on top of it.
//
// Routes
//   POST/GET/HEAD /webhooks/broadridge/:token   <- register this URL in the Toolkit
//   GET  /events?token=ADMIN_TOKEN[&limit=50]   -> newest events (summaries)
//   GET  /events/:id?token=ADMIN_TOKEN          -> one event, full payload
//   GET  /health                                -> liveness
//
// Environment
//   PORT            Railway sets this
//   DATA_DIR        where events are written (mount a Railway volume here), default ./data
//   WEBHOOK_TOKEN   secret path segment Broadridge posts to (required in production)
//   ADMIN_TOKEN     secret for reading captured events (required in production)

import express from "express";
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.resolve("./data");
const EVENTS_DIR = path.join(DATA_DIR, "events");
const LOG_FILE = path.join(DATA_DIR, "events.jsonl");
const WEBHOOK_TOKEN = process.env.WEBHOOK_TOKEN || "";
const ADMIN_TOKEN = process.env.ADMIN_TOKEN || "";
const MAX_BODY = "10mb";

if (!WEBHOOK_TOKEN || !ADMIN_TOKEN) {
  console.warn("[warn] WEBHOOK_TOKEN and/or ADMIN_TOKEN not set — endpoints are unprotected. Set both in production.");
}

fs.mkdirSync(EVENTS_DIR, { recursive: true });

const app = express();
app.disable("x-powered-by");
app.set("trust proxy", true);

// Capture the raw body for every content type so nothing is lost if the
// Toolkit posts form-encoded, XML or signed payloads.
app.use(express.raw({ type: () => true, limit: MAX_BODY }));

function safeEqual(a, b) {
  const ab = Buffer.from(String(a));
  const bb = Buffer.from(String(b));
  return ab.length === bb.length && crypto.timingSafeEqual(ab, bb);
}

function webhookAuthorized(req) {
  if (!WEBHOOK_TOKEN) return true;
  const supplied = req.params.token || req.query.token || "";
  return safeEqual(supplied, WEBHOOK_TOKEN);
}

function adminAuthorized(req) {
  if (!ADMIN_TOKEN) return true;
  const header = req.get("authorization") || "";
  const bearer = header.startsWith("Bearer ") ? header.slice(7) : "";
  const supplied = bearer || req.query.token || "";
  return safeEqual(supplied, ADMIN_TOKEN);
}

function parseBody(raw, contentType) {
  if (!raw || raw.length === 0) return { parsed: null, text: "" };
  const text = raw.toString("utf8");
  const ct = (contentType || "").toLowerCase();
  if (ct.includes("json") || text.trim().startsWith("{") || text.trim().startsWith("[")) {
    try {
      return { parsed: JSON.parse(text), text };
    } catch {
      /* fall through */
    }
  }
  if (ct.includes("x-www-form-urlencoded")) {
    return { parsed: Object.fromEntries(new URLSearchParams(text)), text };
  }
  return { parsed: null, text };
}

// Best-effort guess at the event name so the review list is scannable.
// We don't yet know Broadridge's field names; this looks at the usual suspects.
function guessEventName(parsed, headers) {
  const candidates = ["eventType", "EventType", "event", "Event", "eventName", "EventName", "type", "Type", "topic", "Topic"];
  if (parsed && typeof parsed === "object") {
    for (const key of candidates) {
      if (typeof parsed[key] === "string") return parsed[key];
    }
    for (const container of ["data", "Data", "payload", "Payload"]) {
      const inner = parsed[container];
      if (inner && typeof inner === "object") {
        for (const key of candidates) {
          if (typeof inner[key] === "string") return inner[key];
        }
      }
    }
  }
  for (const h of Object.keys(headers)) {
    if (/event/i.test(h) && typeof headers[h] === "string") return headers[h];
  }
  return null;
}

function recordEvent(req) {
  const id = `${Date.now()}-${crypto.randomBytes(4).toString("hex")}`;
  const contentType = req.get("content-type") || "";
  const { parsed, text } = parseBody(req.body, contentType);
  const record = {
    id,
    receivedAt: new Date().toISOString(),
    method: req.method,
    path: req.path.replace(req.params.token || "", "<token>"),
    ip: req.ip,
    headers: req.headers,
    query: { ...req.query, token: req.query.token ? "<token>" : undefined },
    contentType,
    bodyBytes: req.body ? req.body.length : 0,
    eventName: guessEventName(parsed, req.headers),
    body: parsed ?? text,
  };
  fs.writeFileSync(path.join(EVENTS_DIR, `${id}.json`), JSON.stringify(record, null, 2));
  fs.appendFileSync(LOG_FILE, JSON.stringify(record) + "\n");
  console.log(`[event] ${record.receivedAt} ${record.method} event=${record.eventName ?? "?"} bytes=${record.bodyBytes} id=${id}`);
  return record;
}

// ---------------------------------------------------------------- webhook ---
app.all("/webhooks/broadridge/:token?", (req, res) => {
  if (!webhookAuthorized(req)) {
    console.warn(`[deny] webhook from ${req.ip} with bad token`);
    return res.status(404).end();
  }
  // Some providers probe with GET/HEAD/OPTIONS when an endpoint is registered.
  if (req.method === "HEAD" || req.method === "OPTIONS") return res.status(200).end();
  const record = recordEvent(req);
  res.status(200).json({ ok: true, id: record.id });
});

// ----------------------------------------------------------------- review ---
app.get("/events", (req, res) => {
  if (!adminAuthorized(req)) return res.status(401).json({ error: "unauthorized" });
  const limit = Math.min(parseInt(req.query.limit, 10) || 50, 500);
  const files = fs
    .readdirSync(EVENTS_DIR)
    .filter((f) => f.endsWith(".json"))
    .sort()
    .reverse()
    .slice(0, limit);
  const events = files.map((f) => {
    const r = JSON.parse(fs.readFileSync(path.join(EVENTS_DIR, f), "utf8"));
    return { id: r.id, receivedAt: r.receivedAt, method: r.method, eventName: r.eventName, contentType: r.contentType, bodyBytes: r.bodyBytes };
  });
  res.json({ count: events.length, events });
});

app.get("/events/:id", (req, res) => {
  if (!adminAuthorized(req)) return res.status(401).json({ error: "unauthorized" });
  const file = path.join(EVENTS_DIR, `${path.basename(req.params.id)}.json`);
  if (!fs.existsSync(file)) return res.status(404).json({ error: "not found" });
  res.type("json").send(fs.readFileSync(file, "utf8"));
});

app.get("/health", (_req, res) => {
  res.json({ ok: true, service: "bca-fund-monitor", phase: "webhook-receiver", dataDir: DATA_DIR, uptimeSec: Math.round(process.uptime()) });
});

app.get("/", (_req, res) => res.type("text").send("bca-fund-monitor: webhook receiver online\n"));

app.listen(PORT, () => {
  console.log(`bca-fund-monitor listening on :${PORT}, data in ${DATA_DIR}`);
});
