# Build guide

End to end, from an empty workbook to a narrative arriving in Teams every Monday.

Work through the phases in order — each one ends with a check you can actually
run, so a failure is caught where it happened rather than three phases later.

| Phase | What | Time |
|---|---|---|
| 0 | Accounts, licences and one decision | 15 min |
| 1 | Office Script in Excel | 30 min |
| 2 | Supabase database | 10 min |
| 3 | Digest bot running locally | 20 min |
| 4 | Git repo on GitHub | 10 min |
| 5 | Deploy to Render | 15 min |
| 6 | Power Automate flow | 30 min |
| 7 | End-to-end verification | 15 min |

Phases 1 and 2 are independent — the Office Script is useful on its own, and you
can stop after Phase 1 if you only want the alert worksheets.

---

## Phase 0 — Before you start

### Accounts and access

- **Microsoft 365 business account** with Excel on the web. Open any workbook and
  look for the **Automate** tab. No Automate tab means Office Scripts is not
  licensed or is disabled by your tenant admin — nothing in Phase 1 will work
  until that is sorted.
- **The `DemandAlertsScripts` workbook** stored in SharePoint or OneDrive (not on
  a local disk — Office Scripts and Power Automate both need a cloud file), with
  the IBP extract already in a worksheet named `DBAlerts`.
- **Power Automate** with the *Excel Online (Business)* and *Microsoft Teams*
  connectors.
- **A Teams team and channel** for the digest, where you are allowed to post.
- **Groq** free account → https://console.groq.com/keys
- **Supabase** free account → https://supabase.com
- **GitHub** and **Render** accounts.
- Locally: **Python 3.12+** and **git**.

### The one decision: which Supabase project

The digest needs two tables and one function. You can either:

- **Reuse your existing free Supabase project** (the one the RAG bot uses). The
  table names don't collide with anything there. Simplest, no new keys, no extra
  project against the free-tier limit. **Recommended.**
- **Create a new project.** Cleaner separation, but it counts against the free
  tier's project limit, and your main organisation is on a paid plan where a new
  project is billable — so create it in the *free* organisation if you go this way.

Either way the digest bot gets its own tables and never touches the RAG tables.

---

## Phase 1 — Office Script

### 1.1 Install the script

1. Open the `DemandAlertsScripts` workbook in **Excel on the web**.
2. **Automate → New Script**. The editor opens with a stub `main`.
3. Select everything in the editor and delete it.
4. Open [office-script/src/DemandAlerts.ts](office-script/src/DemandAlerts.ts),
   copy the **entire** file, paste it in.
5. Rename the script to **`Demand Alerts`** (click the name at the top). The
   Power Automate step in Phase 6 looks it up by this exact name.
6. **Save**.

> Do **not** paste `office-script/types/excelscript.d.ts`. That file exists only
> so the script type-checks in a local editor. Excel provides those types itself,
> and pasting it will break the script.

### 1.2 First run — read the log, don't skip this

Click **Run**. Then open the **output/console pane** in the editor.

**This is the most important step in the whole build.** The workbook was never
inspected while the script was written, so the dimension mapping is resolved at
runtime and this log is the only confirmation that it resolved *correctly*.

You are looking for seven lines like:

```text
Dimension 'market' -> column A (0) via configured column index (header blank)
Dimension 'product' -> column B (1) via header name 'Product ID'
Dimension 'productDesc' -> column C (2) via header name 'Product Description'
Dimension 'productLocation' -> column D (3) via configured column index (header blank)
Dimension 'customer' -> column E (4) via header name 'Customer'
Dimension 'customerHierarchy' -> column F (5) via configured column index (header blank)
Dimension 'week' -> column G (6) via header name 'Week'
Current week Monday: 27/07/2026
Reading 48213 data row(s) x 17 column(s) in blocks of 5882 rows.
Source rows read: 48213 (in 9 block(s)) | invalid week: 0 | outside horizon / current week: 9044 | ...
```

Check every line against what is actually in those columns in `DBAlerts`.

### 1.3 If the mapping is wrong

Edit the constants near the top of the script:

```ts
const DIM_COL: { [key: string]: number } = {
    market: 0,            // A
    product: 1,           // B
    productDesc: 2,       // C
    productLocation: 3,   // D
    customer: 4,          // E
    customerHierarchy: 5, // F
    week: 6               // G
};
```

Indexes are 0-based and absolute: `A = 0`, `B = 1`, … `G = 6`. Set a dimension to
`-1` if your extract genuinely does not have it (only `productDesc`,
`productLocation` and `customerHierarchy` may be absent).

If name detection is picking the wrong column, turn it off and rely purely on the
indexes:

```ts
const USE_HEADER_NAME_DETECTION: boolean = false;
```

Run again and re-read the log. Repeat until all seven lines are right.

### 1.4 If it stops with an error

The script refuses to guess. Three errors you may see:

| Error | Fix |
|---|---|
| `Range getValues: the response payload size has exceeded the limit` | The extract is too big to read in one go. The script already reads it in blocks and halves them automatically on failure; if it still fails, lower `READ_CHUNK_CELLS` (default `100000`) near the top of the script. |
| `Missing required Key Figure column(s)` | The message lists what is missing *and* every Key Figure header it did find. Compare them — usually a naming difference in the extract. Correct the `KF_*` constants to match your extract exactly. |
| `The Week / Time Period dimension could not be identified` | Set `DIM_COL.week` to the right index. The column must contain values like `W01 26 (29/12)` with the date in parentheses. |
| `Required dimension(s) could not be identified` | Set the named dimensions in `DIM_COL`. |

### 1.5 Set the volume thresholds

**Do this before anyone else sees the output.** These three constants default to
`100`, which is a placeholder — the volume unit of your extract (cases, hectolitres,
eaches) was unknown when the script was written:

```ts
const ALERT1_MIN_ABS_DIFF: number = 100;          // min forecast change to alert on
const ALERT3_MIN_HISTORICAL_VOLUME: number = 100; // min sales to judge a forecast
const ALERT3_MIN_FUTURE_VOLUME: number = 100;     // min forecast when there is no history
```

Set them to real materiality for your business. Too low and the first run buries
planners in noise; too high and it stays silent.

The percentage thresholds (10%, −5pp, 15%, +5pp) are already per specification —
leave them unless you have a reason.

### 1.6 Check the output

After a clean run you should have four new worksheets:

| Worksheet | Table name |
|---|---|
| `Alert 1 FDP Change` | `tblAlert1FDPChange` |
| `Alert 2 Accuracy Bias` | `tblAlert2AccuracyBias` |
| `Alert 3 Forecast vs Sales` | `tblAlert3ForecastVsSales` |
| `Alert 4 Stat vs FDP` | `tblAlert4StatVsFDP` |

The table names matter — Phase 6 reads them by name.

Sanity-check a few rows against IBP by hand. In particular pick one
Product-Customer row from Alert 2 and confirm the accuracy looks right for that
week, because everything downstream trusts these numbers.

An alert that matched nothing still gets its worksheet **and its table**, holding
one placeholder row reading `No alerts this week` with no Market or Product. That
is deliberate: Phase 6 reads each table by name, and a missing table would fail
the whole weekly run on a quiet week.

> ✅ **Checkpoint.** Four worksheets, four tables, mapping log correct, thresholds
> set, spot-check passed.

---

## Phase 2 — Supabase

### 2.1 Run the migration

1. Open your Supabase project → **SQL Editor** → **New query**.
2. Paste the whole of
   [digest-bot/sql/001_demand_alerts.sql](digest-bot/sql/001_demand_alerts.sql).
3. **Run**.

### 2.2 Verify

In **Table Editor** you should now see `demand_alert_snapshots` and
`demand_digests`. Then run this in the SQL editor:

```sql
select demand_alert_recurrence('alert1', 'Product', current_date - 60);
```

It should return an empty result — not an error. An error means the function did
not get created.

### 2.3 Collect the credentials

**Project Settings → API**:

- `SUPABASE_URL` — the project URL, `https://<ref>.supabase.co`
- `SUPABASE_KEY` — the **secret / service-role** key

> Use the secret key, not the publishable one. Both tables have row level security
> enabled, so the publishable key cannot write to them and every snapshot would
> silently fail.

> ✅ **Checkpoint.** Two tables, one function, both keys copied somewhere safe.

---

## Phase 3 — Digest bot, locally

Prove it works on your machine before deploying it.

### 3.1 Install

```powershell
cd C:\Users\PENPX003\demand-alerts\digest-bot
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 3.2 Generate the endpoint token

```powershell
.\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_urlsafe(32))"
```

Copy the result. This is the shared secret between Power Automate and the bot.

### 3.3 Configure

```powershell
Copy-Item .env.example .env
notepad .env
```

Fill in `GROQ_API_KEY`, `SUPABASE_URL`, `SUPABASE_KEY` and `DEMAND_DIGEST_TOKEN`
(the token from 3.2). Leave the `DEMAND_*` tuning values as they are.

### 3.4 Run the self-test

```powershell
.\.venv\Scripts\python.exe tests\test_digest.py
```

38 checks, no network needed. It ends with `all checks passed`. This covers the
numeric parsing, the aggregation rules and the endpoint auth — if something here
fails, stop and fix it before going further.

### 3.5 Start the service

```powershell
.\.venv\Scripts\python.exe -m uvicorn app.main:app --reload
```

In a second terminal:

```powershell
$t = "<your DEMAND_DIGEST_TOKEN>"
Invoke-RestMethod -Uri http://127.0.0.1:8000/health

Invoke-RestMethod -Uri http://127.0.0.1:8000/api/demand-digest -Method Post `
  -ContentType "application/json" -Headers @{ "X-Digest-Token" = $t } `
  -Body '{"alerts":{"alert1":[],"alert2":[],"alert3":[],"alert4":[]}}'
```

The empty payload returns the no-alerts narrative **without calling the LLM**, so
it proves the wiring, the token and the Supabase connection in one request. If
Supabase is configured correctly you will *not* see
`snapshot/trend step failed` in the service log.

> ✅ **Checkpoint.** Self-test green, `/health` responds, empty digest returns a
> narrative, no Supabase warning in the log.

---

## Phase 4 — Git and GitHub

```powershell
cd C:\Users\PENPX003\demand-alerts
git init
git add .
git commit -m "Demand Planning alerts: Office Script + weekly digest bot"
```

`.gitignore` already excludes `.env` and `.venv`. Confirm before pushing:

```powershell
git status --short          # .env must NOT appear
git ls-files | Select-String "\.env$"   # must return nothing
```

Create an empty repo on GitHub (for example `demand-alerts`), then:

```powershell
git remote add origin https://github.com/<you>/demand-alerts.git
git branch -M main
git push -u origin main
```

> ✅ **Checkpoint.** Repo pushed, `.env` not in it.

---

## Phase 5 — Deploy to Render

The repo holds both halves, so point Render at the `digest-bot` subfolder.

1. Render → **New + → Web Service**.
2. Connect the GitHub repo.
3. Configure:
   - **Root Directory**: `digest-bot`
   - **Runtime**: Docker (it finds `digest-bot/Dockerfile`)
   - **Instance type**: Free
   - **Health check path**: `/health`
4. Add the environment variables:

   | Key | Value |
   |---|---|
   | `GROQ_API_KEY` | your Groq key |
   | `GROQ_CHAT_MODEL` | `llama-3.3-70b-versatile` |
   | `SUPABASE_URL` | from Phase 2.3 |
   | `SUPABASE_KEY` | the secret key from Phase 2.3 |
   | `DEMAND_DIGEST_TOKEN` | the token from Phase 3.2 |
   | `DEMAND_TREND_WEEKS` | `8` |
   | `DEMAND_SNAPSHOT_PC_LIMIT` | `500` |
   | `DEMAND_TOP_N` | `8` |

5. **Create Web Service** and wait for the build.

Verify against the live URL:

```powershell
Invoke-RestMethod -Uri https://<your-service>.onrender.com/health
```

> **Cold start.** Free instances sleep after ~15 minutes idle, and the next
> request takes around 30 seconds. That is harmless for a weekly scheduled flow —
> Power Automate's HTTP timeout absorbs it. Only add an uptime pinger on `/health`
> if you also want the endpoint responsive on demand.

`digest-bot/render.yaml` is a Blueprint equivalent of the above, useful if you
ever split `digest-bot` into its own repo. The manual Web Service path above is
the simpler route for this monorepo.

> ✅ **Checkpoint.** `/health` returns `{"status":"ok"}` from the public URL.

---

## Phase 6 — Power Automate flow

### 6.1 Create the flow

**Create → Scheduled cloud flow**. Name it `Weekly Demand Digest`. Start on a
**Monday**, repeat every **1 week**.

Pick a time *after* your IBP extract lands in `DBAlerts` — the flow refreshes the
alert worksheets but cannot refresh the extract itself. If `DBAlerts` is stale,
the digest confidently describes stale alerts.

### 6.2 Step 2 — Run script

Add **Excel Online (Business) → Run script**.

- **Location / Document Library / File**: the `DemandAlertsScripts` workbook
- **Script**: `Demand Alerts`

This rebuilds the four worksheets, so the digest never reads last week's output.

### 6.3 Steps 3–6 — read the four tables

Add **Excel Online (Business) → List rows present in a table** four times, same
workbook each time, one per table:

| Step | Table |
|---|---|
| 3 | `tblAlert1FDPChange` |
| 4 | `tblAlert2AccuracyBias` |
| 5 | `tblAlert3ForecastVsSales` |
| 6 | `tblAlert4StatVsFDP` |

**On every one of these four actions**, open the **⋯ menu → Settings** and turn
**Pagination** *on*, with **Threshold** set to something comfortably above your
row count (e.g. `5000`).

> This is not optional. Without pagination the action silently returns only the
> first **256** rows, and the digest will describe a fraction of your alerts while
> looking completely healthy. It is the easiest thing in this build to get wrong
> and the hardest to notice.

### 6.4 Step 7 — HTTP

Add an **HTTP** action.

- **Method**: `POST`
- **URI**: `https://<your-service>.onrender.com/api/demand-digest`
- **Headers**:

  | Key | Value |
  |---|---|
  | `Content-Type` | `application/json` |
  | `X-Digest-Token` | your `DEMAND_DIGEST_TOKEN` |

- **Body**:

  ```json
  {
    "week_of": "@{formatDateTime(utcNow(),'yyyy-MM-dd')}",
    "alerts": {
      "alert1": @{body('List_rows_present_in_a_table')?['value']},
      "alert2": @{body('List_rows_present_in_a_table_2')?['value']},
      "alert3": @{body('List_rows_present_in_a_table_3')?['value']},
      "alert4": @{body('List_rows_present_in_a_table_4')?['value']}
    }
  }
  ```

Check the action names against your own flow — Power Automate suffixes duplicated
actions `_2`, `_3`, `_4` in creation order, and if two of these point at the same
table you will get a plausible-looking digest describing the wrong alert.

`week_of` is snapped back to the Monday of that week, so any date works.

### 6.5 Step 8 — post to Teams

Add **Microsoft Teams → Post message in a chat or channel**.

- **Post as**: Flow bot
- **Post in**: Channel
- **Team / Channel**: your Demand Planning team and channel
- **Message**: `body('HTTP')?['narrative_html']`

> Use **`narrative_html`**, not `narrative`. The Teams action treats the message
> as HTML, so the markdown version would post literal `**` around every heading.
> `narrative_html` is the same text with `<b>`, `<p>` and `<ul>` applied.

### 6.6 Test it

**Save**, then **Test → Manually → Run flow**. Watch each step go green and check
the message arrives in the channel.

> ✅ **Checkpoint.** A formatted narrative in your Teams channel.

---

## Phase 7 — End-to-end verification

Run the flow once more and confirm all of this:

- [ ] The four alert worksheets were rebuilt (check a timestamp or a value you changed).
- [ ] Each `List rows` step returned the row count you expect — **not exactly 256**.
- [ ] The HTTP step returned `200`.
- [ ] `total_alert_rows` in the response matches the sum of the four table row counts.
- [ ] `snapshot_rows_stored` is greater than zero.
- [ ] `demand_alert_snapshots` in Supabase has rows for this week.
- [ ] `demand_digests` has one row for this week.
- [ ] The Teams message is formatted, with no literal `**`.
- [ ] Pick three figures from the narrative and find them in the alert worksheets.

That last check is the one that matters. Do it properly on the first run.

### Confirming the trend feature

Recurrence only appears from the **second** week onward — the first run has
nothing to compare against, and the narrative will say so. After two runs, look
for phrases like *"alerted in 2 of the last 8 weeks"*. If week two still reports
nothing recurring, check `demand_alert_snapshots` actually has two distinct
`week_of` values.

---

## Operating it

### Weekly

Nothing. The flow runs itself. Glance at the Teams post.

### When a narrative looks wrong

Re-send the same payload with `"include_brief": true` in the HTTP body. The
response then contains the exact brief the model was given.

- **Brief is right, narrative is wrong** → the model misread it. Tighten `PROMPT`
  in `digest-bot/digest/digest.py`.
- **Brief is wrong** → the bug is upstream, in `digest-bot/digest/alerts.py` or in
  the Office Script. Trace the figure back to the worksheet.

The model is instructed to use only figures from the brief, but instructions are
not guarantees. The spot-check in Phase 7 is what catches a drifting model.

### Troubleshooting

| Symptom | Cause |
|---|---|
| `401 Unauthorized` | `X-Digest-Token` missing or not matching `DEMAND_DIGEST_TOKEN` on Render. |
| `List rows` step fails, table not found | The Office Script did not run, or a table was renamed. Table names are in Phase 1.6. |
| Exactly 256 rows per alert | Pagination is off. See 6.3. |
| Narrative never reports anything recurring | Migration not applied, or wrong Supabase key. Check the Render log for `snapshot/trend step failed`. |
| `snapshot_rows_stored: 0` | Same cause. |
| Literal `**` in the Teams message | Step 8 is using `narrative` instead of `narrative_html`. |
| Digest describes last week's data | The flow ran before the IBP extract refreshed. |
| Flow times out | Render cold start; raise the HTTP action's timeout. |
| `Digest failed: GROQ_API_KEY is not set` | Env var missing on Render. |

### Tuning

| Want | Change |
|---|---|
| Fewer / more alerts | The threshold constants in the **Office Script**. The digest narrates whatever the alerts contain. |
| More items named in the narrative | `DEMAND_TOP_N` on Render. |
| Longer trend memory | `DEMAND_TREND_WEEKS` on Render. |
| Different tone or sections | `PROMPT` in `digest-bot/digest/digest.py`. |

After changing anything in `digest-bot`, run `tests\test_digest.py` and push —
Render redeploys on push to `main`.

---

## What is not covered

- **The Office Script has never run against real data.** Phase 1.2 is where you
  find out whether the dimension mapping is right.
- **The volume thresholds are placeholders.** Phase 1.5.
- **No alerting on failure.** If the flow breaks, nothing tells you except the
  absent Teams post. Add a failure notification in Power Automate if this becomes
  business-critical.
