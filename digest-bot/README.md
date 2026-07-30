# Demand digest bot

FastAPI service that turns the four Demand Planning alert tables into a weekly
narrative for a Teams channel.

Standalone: no shared code, no shared Supabase tables and no shared Teams channel
with any other bot. Same free-tier stack (Groq + Supabase + Power Automate),
separate everything else.

## Why the numbers are safe

This is **not** RAG and there is no retrieval. The model never sees an alert row
and never calculates anything. `digest/alerts.py` computes every figure in Python,
`digest/digest.py` renders them into a factual brief, and the prompt tells the
model to narrate that brief and nothing else.

Two mistakes are prevented by construction rather than by asking the model nicely:

- **Double counting.** Alert tables hold both aggregation levels
  (Product-Customer, and Product with Customer = `ALL`). Totals use Product-level
  rows only.
- **Percent vs percentage points.** Every figure has a declared kind. Levels
  render as `82.0%`, differences as `-21.0 pp`. There is no `%` in the brief for
  the model to copy onto a difference.

Volume totals are labelled per alert, because they mean different things: Alerts 1
and 3 report forecast *change* volume, Alerts 2 and 4 report the *actual sales*
covered. They are not comparable and the brief says so.

## Layout

```
app/main.py            FastAPI endpoints
digest/config.py       env -> Config
digest/alerts.py       payload parsing + deterministic statistics
digest/digest.py       Supabase snapshots, recurrence, prompt, orchestration
digest/llm.py          Groq client with retry
sql/001_demand_alerts.sql
docs/power-automate-flow.md
```

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET | `/health` | Liveness (Render health check, keep-warm pinger) |
| POST | `/api/demand-digest` | Alert tables in, narrative out |
| GET | `/api/demand-digest/latest` | Most recently generated digest |

Both `/api/*` endpoints require the `X-Digest-Token` header when
`DEMAND_DIGEST_TOKEN` is set.

## Setup

### 1. Database

Run `sql/001_demand_alerts.sql` in the Supabase SQL editor. It creates
`demand_alert_snapshots`, `demand_digests` and the `demand_alert_recurrence` RPC.

Without it the digest still works — the snapshot and trend steps fail soft and log
`snapshot/trend step failed, continuing without history` — but nothing can be
described as recurring, which is the most useful part.

Use the **secret / service-role** Supabase key: the tables have row level security
enabled, so the publishable key cannot write to them.

### 2. Local run

```powershell
py -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
Copy-Item .env.example .env    # then fill it in
.\.venv\Scripts\python.exe -m uvicorn app.main:app --reload
```

Smoke test — an empty payload returns the no-alerts narrative without calling the
LLM, which proves the wiring before you point a flow at it:

```powershell
curl.exe -X POST http://127.0.0.1:8000/api/demand-digest `
  -H "Content-Type: application/json" `
  -d '{\"alerts\":{\"alert1\":[],\"alert2\":[],\"alert3\":[],\"alert4\":[]}}'
```

### 3. Deploy

Push this folder as its own repo, then Render → **New + → Blueprint**. Set the
`sync: false` secrets in the dashboard: `GROQ_API_KEY`, `SUPABASE_URL`,
`SUPABASE_KEY`, `DEMAND_DIGEST_TOKEN`.

Generate the token with:

```
python -c "import secrets; print(secrets.token_urlsafe(32))"
```

> **Keep-warm (optional).** Render free instances sleep after ~15 min idle; the
> first request then takes ~30s. That is harmless for a weekly scheduled flow —
> Power Automate's HTTP timeout absorbs it. Only add an uptime pinger on
> `/health` if you also want the endpoint responsive on demand.

### 4. Power Automate flow

See [../SETUP.md](../SETUP.md) — the full build guide, phase 6.

Post **`narrative_html`** to Teams, not `narrative`: the Teams action renders the
message as HTML, so the markdown version shows literal `**` around every heading.

## Narrative shape

Six fixed sections: **Headline**, **Forecast changes vs last cycle**,
**Accuracy and bias**, **Forecast vs recent sales**,
**Where the statistical forecast is winning**, **Suggested focus this week**.
Around 400 words, capped at 600. Recurring combinations are called out explicitly.

A week with no alerts at all skips the LLM entirely and returns a fixed message.

## Tuning

| Variable | Default | Effect |
|---|---|---|
| `DEMAND_TREND_WEEKS` | `8` | How far back "recurring" looks |
| `DEMAND_SNAPSHOT_PC_LIMIT` | `500` | Product-Customer rows stored per alert per week |
| `DEMAND_TOP_N` | `8` | Items per alert and level in the brief |
| `GROQ_CHAT_MODEL` | `llama-3.3-70b-versatile` | Generation model |
| `WORKBOOK_URL` | *(blank)* | Adds "Open the alert workbook" to the digest. http(s) only; blank omits it |

Too noisy? Raise the thresholds in the **Office Script**, not here — the digest
narrates whatever the alerts contain. Different tone or sections? Edit `PROMPT` in
`digest/digest.py`.
