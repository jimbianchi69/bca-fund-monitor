# bca-fund-monitor

BCA Financial Group's quarterly 401(k) fund-monitoring service. Receives events from the
Broadridge Fiduciary Focus Toolkit™ and (in later phases) assembles the quarterly
3(21)/3(38) monitoring package and keeps Salesforce current.

## Phase 1 — webhook receiver (this version)

Discovery only: capture every Toolkit webhook untouched so we can learn the real event
schema before building on it.

### Routes

| Route | Purpose |
|---|---|
| `POST /webhooks/broadridge/<WEBHOOK_TOKEN>` | Register this URL in Toolkit → Workspace Management → Webhooks |
| `GET /events?token=<ADMIN_TOKEN>` | Newest captured events (summaries) |
| `GET /events/<id>?token=<ADMIN_TOKEN>` | One event, full headers + body |
| `GET /health` | Liveness check |

### Environment variables

| Name | Purpose |
|---|---|
| `WEBHOOK_TOKEN` | Secret path segment in the webhook URL. Anything posted without it gets a 404. |
| `ADMIN_TOKEN` | Secret for reading captured events. |
| `DATA_DIR` | Where events are persisted. On Railway, mount a volume at `/data` and set this to `/data`. |
| `PORT` | Set by Railway. |

### Toolkit events to subscribe

- **Toolkit Data Update: New Quarterly Data Available** — quarter's starting gun
- **Toolkit Monitoring Report: ClientReadyGenerated** — finished report per plan
- **Toolkit Monitoring Report: Completed** — belt-and-suspenders

### Run locally

```bash
npm install
WEBHOOK_TOKEN=dev ADMIN_TOKEN=dev npm start
curl -X POST localhost:3000/webhooks/broadridge/dev -H 'content-type: application/json' -d '{"eventType":"test"}'
curl 'localhost:3000/events?token=dev'
```

## Roadmap

- **Phase 2** — on `ClientReadyGenerated`, fetch the report/data from Broadridge and store it per plan
- **Phase 3** — quarterly run (Claude scheduled task + MCP tools): sponsor package, watch-list memo, 3(38) decision memo, minutes template; file to client folder; update the 401(k) policy record and log activity in Salesforce via bca-salesforce-mcp
- **Phase 4** — expected-vs-received fee reconciliation against commission statement items

## Repo layout

| Path | What |
|---|---|
| `server.js` | Phase 1 webhook receiver (Railway) |
| `tools/build_screen.py` | Builds the quarterly screening workbook from the Fidelity FPRS platform list |
| `docs/QUARTERLY_RUNBOOK.md` | The quarterly monitoring procedure (mirrors the `bca-quarterly-fund-monitoring` skill) |
