# Build guide

End to end, from an empty workbook to a narrative arriving in Teams every Monday.

Work through the phases in order — each ends with a check you can actually run, so
a failure surfaces where it happened rather than three phases later.

| Phase | What | Time |
|---|---|---|
| 0 | Accounts, licences and one decision | 15 min |
| 1 | Office Script in Excel | 30 min |
| 2 | Supabase database | 10 min |
| 3 | Digest bot running locally | 20 min |
| 4 | Git repo on GitHub | 10 min |
| 5 | Deploy to Render | 15 min |
| 6 | Power Automate flow | 40 min |
| 7 | End-to-end verification | 15 min |

Phases 1 and 2 are independent. The Office Script is useful on its own — stop after
Phase 1 if you only want the alert worksheets in Excel.

**This deployment** (fill in your own if rebuilding elsewhere):

| Thing | Value |
|---|---|
| Workbook | `/Ai Agents/DBAlerts.xlsx` in the Teams channel *Ai Agents* |
| Source worksheet | `DBAlerts` |
| Office Script name | `Demand Alerts` |
| Repo | `github.com/penpx003/demand-alerts` |
| Render service | `demand-alerts` → `https://demand-alerts.onrender.com` |
| Flow | `Weekly Demand Digest` |

---

## Known traps

Everything here cost real debugging time on the first build. If you are rebuilding,
read this table first — it is the whole hard-won lesson of the project in one place.

| Trap | Symptom | Where |
|---|---|---|
| Office Scripts payload limit on **read** | `Range getValues: the response payload size has exceeded the limit` | 1.4 |
| Payload limit on **write** — hits Alert 1 first, and it publishes first, so alerts 2–4 keep *stale* output | Worksheet exists with **no table**; script appears fine | 1.4 |
| Excel connector serves a **cached** workbook | `Run script` Succeeded, next step fails `No table was found` | 6.3 |
| Table bound by **GUID**, and the script recreates tables each run | Works once, then `No table was found with the name '{GUID}'` every week | 6.4 |
| Pagination **off** by default | Silently only the first **256** rows; digest looks healthy but describes a fraction | 6.4 |
| `alertN` keys **crossed** in the HTTP body | Plausible digest describing the wrong alert; `metric`/`volume` null | 6.5 |
| Teams message **typed** rather than entered as an expression | Channel receives the literal `body('HTTP')?['narrative_html']` | 6.6 |
| PowerShell `echo x >> file` writes **UTF-16** | git commits README.md as a binary blob | 4 |
| `git remote add` when origin exists | Fails, silently leaves the **wrong** URL; push dies with bare `error: 400` | 4 |
| Country flow not sending `country` | Countries **overwrite each other's** digest; trends mix countries | 8 |

---

## Phase 0 — Before you start

### Accounts and access

- **Microsoft 365 business account** with Excel on the web. Open any workbook and
  look for the **Automate** tab. No Automate tab means Office Scripts is not
  licensed or is disabled by your tenant admin — nothing in Phase 1 works until
  that is sorted.
- **The workbook** stored in SharePoint or OneDrive — not on a local disk, as both
  Office Scripts and Power Automate need a cloud file — with the IBP extract in a
  worksheet named `DBAlerts`.
- **Power Automate** with the *Excel Online (Business)* and *Microsoft Teams*
  connectors.
- **A Teams team and channel** for the digest, where you are allowed to post.
- **Groq** free account → <https://console.groq.com/keys>
- **Supabase** free account → <https://supabase.com>
- **GitHub** and **Render** accounts.
- Locally: **Python 3.12+** and **git**.

### The one decision: which Supabase project

The digest needs two tables and one function.

- **Reuse an existing free Supabase project.** Table names don't collide with
  anything else. No new keys, no extra project against the free-tier limit.
  **Recommended.**
- **Create a new project.** Cleaner separation, but it counts against the free
  tier's project limit — and if your main organisation is on a paid plan, a new
  project there is billable. Create it in a *free* organisation.

Either way the digest gets its own tables and touches nothing else.

---

## Phase 1 — Office Script

### 1.1 Install the script

1. Open the workbook in **Excel on the web**.
2. **Automate → New Script**. The editor opens with a stub `main`.
3. Select everything in the editor and delete it.
4. Open [office-script/src/DemandAlerts.ts](office-script/src/DemandAlerts.ts),
   copy the **entire** file, paste it in.
5. Rename the script to **`Demand Alerts`** — click the name at the top. Phase 6
   looks it up by this exact name.
6. **Save**.

> Do **not** paste `office-script/types/excelscript.d.ts`. That file exists only so
> the script type-checks in a local editor. Excel provides those types itself, and
> pasting it breaks the script.

Also install the diagnostic while you are here: **Automate → New Script**, paste
[office-script/src/ListTables.ts](office-script/src/ListTables.ts), name it
`List Tables`. It prints every worksheet and table in the workbook, and is the
fastest way to answer "does that table actually exist?" later.

### 1.2 First run — read the log

Click **Run**, then open the **output/console pane**.

**This is the most important step in the build.** The dimension mapping is resolved
at runtime from your header row, and this log is the only confirmation it resolved
*correctly*. A wrong-but-confident mapping produces plausible nonsense downstream.

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
Alert 1 FDP Change: 312 row(s) — 254 Product-Customer, 58 Product.
...
All four output tables verified present.
```

Check every `Dimension` line against what is really in that column. Then check:

- **`invalid week`** should be near zero. A large number means the week labels are
  not parsing and rows are being dropped silently.
- **The four `Alert N` counts.** All zero is a red flag, not a clean bill of health
  — see 1.5.
- **`All four output tables verified present.`** must be the last line. If the run
  throws `Output incomplete`, some worksheets still hold the *previous* run's data.

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

Indexes are 0-based and absolute: `A = 0` … `G = 6`. Set a dimension to `-1` if
your extract genuinely lacks it — only `productDesc`, `productLocation` and
`customerHierarchy` may be absent.

If name detection picks the wrong column, disable it and rely on the indexes:

```ts
const USE_HEADER_NAME_DETECTION: boolean = false;
```

Re-run and re-read the log until all seven lines are right.

### 1.4 If it stops with an error

The script refuses to guess, so its errors are specific:

| Error | Fix |
|---|---|
| `Range getValues: the response payload size has exceeded the limit` | The extract is too large for one read. The script already reads in blocks and halves them on failure; if it still fails, lower `READ_CHUNK_CELLS` (default `100000`). |
| `Could not write output rows … even at the minimum block size` | Same limit, writing. Lower `WRITE_CHUNK_CELLS`. |
| `Output incomplete — these tables were not created: …` | A publish step failed. Other worksheets may hold stale data from the previous run. Re-run; if it repeats, lower `WRITE_CHUNK_CELLS`. |
| `Missing required Key Figure column(s)` | The message lists what is missing **and** every Key Figure header it found. Usually a naming difference — correct the `KF_*` constants to match your extract exactly. |
| `The Week / Time Period dimension could not be identified` | Set `DIM_COL.week`. The column must hold values like `W01 26 (29/12)` with the date in parentheses. |
| `Required dimension(s) could not be identified` | Set the named dimensions in `DIM_COL`. |

### 1.5 Set the volume thresholds

**Do this before anyone else sees the output.** These three default to `100`, a
placeholder chosen without knowing your volume unit (cases, hectolitres, eaches):

```ts
const ALERT1_MIN_ABS_DIFF: number = 100;          // min forecast change to alert on
const ALERT3_MIN_HISTORICAL_VOLUME: number = 100; // min sales to judge a forecast
const ALERT3_MIN_FUTURE_VOLUME: number = 100;     // min forecast when there is no history
```

Set them to real materiality. Too low buries planners in noise; too high stays
silent. Use the Alert row counts from 1.2 to calibrate — a five-figure Alert 1 count
means the threshold is far too low, and will also push the script towards the
Office Scripts runtime limit when it runs from Power Automate.

The percentage thresholds (10%, −5 pp, 15%, +5 pp) are per specification. Leave them
unless you have a reason.

### 1.6 Check the output

Four worksheets, four tables:

| Worksheet | Table name |
|---|---|
| `Alert 1 FDP Change` | `tblAlert1FDPChange` |
| `Alert 2 Accuracy Bias` | `tblAlert2AccuracyBias` |
| `Alert 3 Forecast vs Sales` | `tblAlert3ForecastVsSales` |
| `Alert 4 Stat vs FDP` | `tblAlert4StatVsFDP` |

Run `List Tables` and confirm all four report **FOUND**. Phase 6 reads them by
these exact names.

An alert that matched nothing still gets its worksheet **and its table**, holding a
single placeholder row reading `No alerts this week` with no Market or Product. That
is deliberate — Phase 6 reads each table by name, and a missing table would fail the
whole weekly run on a quiet week. The digest bot skips that row.

Finally, spot-check by hand. Pick one Product-Customer row from Alert 2 and verify
the accuracy against IBP for that week. Everything downstream trusts these numbers.

> ✅ **Checkpoint.** Mapping log correct, thresholds set, `List Tables` shows 4/4
> FOUND, spot-check passed.

---

## Phase 2 — Supabase

### 2.1 Run the migration

1. Supabase project → **SQL Editor** → **New query**.
2. Paste all of
   [digest-bot/sql/001_demand_alerts.sql](digest-bot/sql/001_demand_alerts.sql).
3. **Run**.
4. Repeat with
   [digest-bot/sql/002_country.sql](digest-bot/sql/002_country.sql), which adds
   the per-country scope. Required even for a single country — the service writes
   a `country` column on every row.

> Migrations run in order and are safe to re-run. If you deploy the service
> without 002, digests still generate and post, but nothing is persisted and no
> trend history accumulates — the log shows
> `column demand_digests.country does not exist`.

### 2.2 Verify

**Table Editor** should now show `demand_alert_snapshots` and `demand_digests`.
Then run:

```sql
select demand_alert_recurrence('alert1', 'Product', current_date - 60);
```

An empty result is correct. An error means the function wasn't created.

### 2.3 Collect the credentials

**Project Settings → API**:

- `SUPABASE_URL` — `https://<ref>.supabase.co`
- `SUPABASE_KEY` — the **secret / service-role** key

> Use the secret key, not the publishable one. Both tables have row level security
> enabled, so the publishable key cannot write and every snapshot fails silently —
> you would get digests forever with no trend history.

> ✅ **Checkpoint.** Two tables, one function, credentials saved.

---

## Phase 3 — Digest bot, locally

Prove it works on your machine before deploying.

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

This is the shared secret between Power Automate and the bot.

### 3.3 Configure

```powershell
Copy-Item .env.example .env
notepad .env
```

Fill in `GROQ_API_KEY`, `SUPABASE_URL`, `SUPABASE_KEY`, `DEMAND_DIGEST_TOKEN` and
optionally `WORKBOOK_URL`. Leave the `DEMAND_*` tuning values alone.

### 3.4 Run the self-test

```powershell
.\.venv\Scripts\python.exe tests\test_digest.py
```

52 checks, no network needed, ending in `all checks passed`. It covers numeric
parsing of Power Automate's formatted cells, the aggregation rules, percent versus
percentage points, mis-wired-table detection, the workbook link and endpoint auth.
If anything fails here, stop and fix it.

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

The empty payload returns the no-alerts narrative **without calling the LLM**, so it
proves the wiring, the token and the Supabase connection in one request. With
Supabase configured you should *not* see `snapshot/trend step failed` in the log.

> ✅ **Checkpoint.** Self-test green, `/health` responds, empty digest returns a
> narrative, no Supabase warning.

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
git status --short                        # .env must NOT appear
git ls-files | Select-String "\.env$"     # must return nothing
```

Create an **empty** repo on GitHub — no README, no .gitignore, no licence. You
already have them, and letting GitHub add its own creates a commit you don't have,
which makes the first push bounce.

> ⚠️ **Ignore GitHub's "…or create a new repository on the command line" snippet.**
> It is written for bash and damages this repo two ways.
> `echo "# demand-alerts" >> README.md` in PowerShell writes **UTF-16**, so git
> stops seeing README.md as text and commits it as a binary blob. And it re-runs
> `git init` + `git commit` over work you have already committed.

Set the remote with `set-url`, which works whether or not one already exists — plain
`git remote add` fails with *"remote origin already exists"* and, worse, leaves the
old wrong URL in place:

```powershell
git remote add origin https://github.com/YOUR-USERNAME/demand-alerts.git 2>$null
git remote set-url origin https://github.com/YOUR-USERNAME/demand-alerts.git
git remote -v          # confirm BOTH lines show your real username
git branch -M main
git push -u origin main
```

Substitute your real username. **Check `git remote -v` before pushing** — a
placeholder left in the URL fails with a bare `error: 400` that gives no hint.

**If the push is rejected** ("fetch first" / non-fast-forward), the GitHub repo has a
commit you don't:

```powershell
git pull --rebase origin main
git push -u origin main
```

**Commit identity.** If git says *"Your name and email address were configured
automatically"* it guessed from your Windows account, and commits won't link to your
GitHub profile unless that email is registered there:

```powershell
git config --global user.name "Your Name"
git config --global user.email "the-email-on-your-github-account"
```

Future commits only.

> ✅ **Checkpoint.** `git remote -v` shows your username, push succeeded, `.env` not
> in the repo, README.md renders as text on GitHub.

---

## Phase 5 — Deploy to Render

The repo holds both halves, so point Render at the `digest-bot` subfolder.

1. Render → **New + → Web Service**.
2. Connect the GitHub repo. If it isn't listed, click **Configure account** and
   grant Render access to it on GitHub.
3. Configure:
   - **Root Directory**: `digest-bot` ← without this the build fails
     `Dockerfile not found`
   - **Runtime**: Docker
   - **Instance Type**: Free
   - **Health Check Path**: `/health`
4. Environment variables:

   | Key | Value |
   |---|---|
   | `GROQ_API_KEY` | your Groq key |
   | `GROQ_CHAT_MODEL` | `openai/gpt-oss-120b` (see note below) |
   | `SUPABASE_URL` | from 2.3 |
   | `SUPABASE_KEY` | the **secret** key from 2.3 |
   | `DEMAND_DIGEST_TOKEN` | the token from 3.2 |
   | `DEMAND_TREND_WEEKS` | `8` |
   | `DEMAND_SNAPSHOT_PC_LIMIT` | `500` |
   | `DEMAND_TOP_N` | `8` |
   | `DEMAND_MAX_OUTPUT_TOKENS` | `4000` |
   | `WORKBOOK_URL` | link to the workbook (below) |

   **`WORKBOOK_URL`** appends *"Open the alert workbook"* to every digest so a
   planner can click from the Teams post to the detail. Get it by opening the
   workbook in Excel on the web and copying the address bar, or via **Copy link** in
   SharePoint. Only `http(s)` URLs are accepted; blank omits the link. The flow can
   override it per run with a `workbook_url` body field.

5. **Create Web Service** and wait for the build.

### Get your real URL

The service page shows the URL under the service name, next to the status badge.
**Copy it from there — do not assume it matches the service name.** `*.onrender.com`
subdomains are globally unique, so a taken name gets a suffix like
`demand-alerts-a7x9.onrender.com`.

Wait for **Live**, then:

```powershell
Invoke-RestMethod -Uri https://YOUR-SERVICE.onrender.com/health
```

Expect `status : ok`.

A `404` carrying the header `x-render-routing: no-server` means nothing is bound to
that hostname — wrong URL, or not Live yet. That is different from your app
returning 404.

> **Cold start.** Free instances sleep after ~15 minutes idle and the next request
> takes ~30 seconds. Harmless for a weekly flow; Power Automate's HTTP timeout
> absorbs it. Add an uptime pinger on `/health` only if you want the endpoint
> responsive on demand.

`digest-bot/render.yaml` is a Blueprint equivalent, useful if you ever split
`digest-bot` into its own repo. Keep its `name:` matching the live service, or a
Blueprint deploy creates a second one.

> ✅ **Checkpoint.** `/health` returns `{"status":"ok"}` from the public URL.

---

## Phase 6 — Power Automate flow

Two rules that apply to the whole flow:

> ⚠️ **All five Excel actions must point at the SAME file.** `Run script` and the
> four `List rows` actions each store their own file reference. If they diverge, the
> script builds tables in one workbook while the flow reads another, and every
> `List rows` fails with "No table was found" even though the tables plainly exist.
> Verify with **⋯ → Peek code** and compare each `metadata` path.
>
> ⚠️ **All four tables must already exist before you can save the flow.** Step 2
> creates them at runtime, but Power Automate validates the `List rows` references
> at **save** time. Run the Office Script by hand first (Phase 1). Saving too early
> fails with `operation 'GetTable' failed with status code 'NotFound'` /
> `ItemNotFound` — which also appears if the **File** reference is stale, since
> Graph's `ItemNotFound` refers to the file as often as the table.

### 6.1 Step 1 — Recurrence

**Create → Scheduled cloud flow**, named `Weekly Demand Digest`. Start **Monday**,
repeat every **1 week**.

Pick a time *after* your IBP extract lands in `DBAlerts`. The flow refreshes the
alert worksheets but cannot refresh the extract — if `DBAlerts` is stale, the digest
confidently describes stale alerts.

### 6.2 Step 2 — Run script

**Excel Online (Business) → Run script**.

- **Location / Document Library / File**: the workbook
- **Script**: `Demand Alerts`

This rebuilds the four worksheets, so the digest never reads last week's output.

### 6.3 Step 3 — Delay (required)

Add **Delay**, **1 minute**.

> Not padding. The Excel connector can serve a cached view of the workbook from
> *before* the script rebuilt the sheets, so the next action fails with
> `No table was found` for a table that cannot be absent — the script throws rather
> than finish without it. The delay lets the rebuilt workbook settle.
>
> Signature symptom: `Run script` reports **Succeeded** and the very next
> `List rows` fails on a missing table. If you see that pair, the answer is the
> delay, not the table.

### 6.4 Steps 4–7 — read the four tables

Add **Excel Online (Business) → List rows present in a table** four times, same
workbook each time. **Rename each action after its table** — it makes the mapping in
6.5 self-evident and the errors readable:

| Action name | Table |
|---|---|
| `List rows present in a table tblAlert1FDPChange` | `tblAlert1FDPChange` |
| `List rows present in a table tblAlert2AccuracyBias` | `tblAlert2AccuracyBias` |
| `List rows present in a table tblAlert3ForecastVsSales` | `tblAlert3ForecastVsSales` |
| `List rows present in a table tblAlert4StatVsFDP` | `tblAlert4StatVsFDP` |

**Type the table name as text. Never pick it from the dropdown.**

- **New designer**: click the field, use the toggle on the right of the dropdown to
  switch to free text (pencil / `fx` icon), type the name.
- **Classic designer**: open the dropdown → **Enter custom value** → type the name.

> ⚠️ The picker stores the table's internal **ID**, a GUID like
> `{2BD3FE1D-C6C1-446F-AA6E-D43731BDD673}`. Step 2 deletes and recreates each
> worksheet every run, so the rebuilt table gets a **new GUID** and the flow keeps
> hunting the old one. It works once, then fails weekly with
> `No table was found with the name '{2BD3FE1D-...}'`. Names survive the rebuild.
>
> Confirm with **Peek code** that `"table"` reads `tblAlert1FDPChange`, not a GUID.
> The **File** field is fine as an item ID — the workbook is never recreated, only
> its sheets.

**On all four actions**: **⋯ → Settings → Pagination on**, **Threshold** comfortably
above your row count (e.g. `5000`).

> Without pagination the action silently returns only the first **256** rows. The
> digest then describes a fraction of your alerts while looking completely healthy.
> The easiest thing here to get wrong and the hardest to notice.

### 6.5 Step 8 — HTTP

Add an **HTTP** action.

- **Method**: `POST`
- **URI**: `https://YOUR-SERVICE.onrender.com/api/demand-digest`
  *This deployment:* `https://demand-alerts.onrender.com/api/demand-digest`
- **Headers**:

  | Key | Value |
  |---|---|
  | `Content-Type` | `application/json` |
  | `X-Digest-Token` | your `DEMAND_DIGEST_TOKEN` |

- **Body** — build it as a single expression. With the actions renamed as in 6.4,
  the mapping reads left to right and is hard to get wrong:

  ```text
  json(concat('{"week_of":"',formatDateTime(utcNow(),'yyyy-MM-dd'),'","country":"ES","workbook_url":"https://YOUR-SITE/sites/Team/Ai%20Agents/DB_Alert.xlsx","alerts":{"alert1":',string(outputs('List_rows_present_in_a_table_tblAlert1FDPChange')?['body/value']),',"alert2":',string(outputs('List_rows_present_in_a_table_tblAlert2AccuracyBias')?['body/value']),',"alert3":',string(outputs('List_rows_present_in_a_table_tblAlert3ForecastVsSales')?['body/value']),',"alert4":',string(outputs('List_rows_present_in_a_table_tblAlert4StatVsFDP')?['body/value']),'}}'))
  ```

  Action references use underscores for spaces. Note the closing `'}}'` — two
  braces, one for `alerts` and one for the outer object. A single brace produces
  malformed JSON.

  `country` and `workbook_url` are what make one flow per country work — see
  [Phase 8](#phase-8--running-several-countries). For a single country, drop both
  and the service falls back to `WORKBOOK_URL` on Render.

> ⚠️ **`alertN` must reference the action reading `tblAlertN…`.** All four tables
> share Market / Product / Customer, so a crossed wire yields a digest that looks
> entirely normal while describing the wrong alert under each heading — the
> alert-specific columns just come back null. The service rejects this
> (`AlertColumnMismatch`), but the check can only fire when rows are present: an
> `alertN` hardcoded to `[]` passes and quietly reports "no alerts this week".

**Alternative body**, if you prefer the raw-JSON form:

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

Power Automate suffixes duplicated actions `_2`, `_3`, `_4` in creation order, so
check the names against your own flow. `week_of` is snapped back to the Monday of
that week, so any date works.

### 6.6 Step 9 — post to Teams

**Microsoft Teams → Post message in a chat or channel**.

- **Post as**: Flow bot
- **Post in**: Channel
- **Team / Channel**: your Demand Planning team and channel
- **Message**: enter as an **expression**:

  ```text
  body('HTTP')?['narrative_html']
  ```

> ⚠️ **Use the Expression tab.** Click the Message box, open the dynamic-content
> panel, switch to **Expression** (`fx`), paste the line, click **OK / Add**. The box
> should then show a coloured token, not raw text. Typed as plain text, the channel
> receives the literal formula instead of the digest.
>
> Use **`narrative_html`**, not `narrative`. Teams renders the message as HTML, so
> the markdown version posts literal `**` around every heading. `narrative_html` is
> the same text with `<b>`, `<p>`, `<ul>` and the workbook link applied.

### 6.7 Test

**Save**, then **Test → Manually → Run flow**. Every step should go green and the
message should arrive in the channel.

> ✅ **Checkpoint.** A formatted narrative in Teams, with a working workbook link.

---

## Phase 7 — End-to-end verification

Green steps are not the same as correct output. Run the flow once more and confirm:

- [ ] The four alert worksheets were rebuilt (check a value you changed by hand).
- [ ] Each `List rows` step returned the count you expect — **not exactly 256**.
- [ ] The HTTP step returned `200`.
- [ ] `total_alert_rows` matches the sum of the four table row counts.
- [ ] `snapshot_rows_stored` is greater than zero.
- [ ] `demand_alert_snapshots` has rows for this week in Supabase.
- [ ] `demand_digests` has one row for this week.
- [ ] The Teams message is formatted, with no literal `**` and no raw expression.
- [ ] The workbook link opens the right file.
- [ ] **Pick three figures from the narrative and find them in the worksheets.**

Then check nothing is silently null — the signature of a crossed `alertN` mapping:

```sql
select alert, agg_level, market, product, metric, volume, alert_type
from demand_alert_snapshots
where week_of = current_date - extract(dow from current_date)::int + 1
limit 20;
```

`market` and `product` populated while `metric` and `volume` are null means an alert
received another alert's rows. Go back to 6.5.

### Confirming the trend feature

Recurrence appears only from the **second** week — the first run has nothing to
compare against and says so. After two runs look for *"alerted in 2 of the last 8
weeks"*. If week two still reports nothing recurring, check
`demand_alert_snapshots` really has two distinct `week_of` values.

---

## Phase 8 — Running several countries

One flow per country, one Teams channel per country, one workbook per country —
but **one** Render service and **one** Supabase project shared by all of them.
Nothing needs duplicating on the hosting side.

```text
ES:  channel AI_Alert (ES)  ->  DB_Alert.xlsx (ES)  ->  flow "Weekly Demand Digest ES"  \
PT:  channel AI_Alert (PT)  ->  DB_Alert.xlsx (PT)  ->  flow "... PT"                    >-- one service
IT:  channel AI_Alert (IT)  ->  DB_Alert.xlsx (IT)  ->  flow "... IT"                   /
```

### What makes it safe

Each POST carries a `country`. It scopes both stored tables and every trend
lookup.

> ⚠️ **Without `country`, countries overwrite each other.** A digest is keyed on
> `(country, week_of)`. Keyed on `week_of` alone — the original design — the last
> country to run each week would replace every other country's digest for that
> week, and "alerted in 3 of the last 8 weeks" would silently count other
> countries' weeks as if they were the same book of business. Migration 002 is
> what fixes this; apply it before adding the second country.

The label is normalised — upper-cased, trimmed, punctuation stripped — so `es`,
`Es` and `ES` are one scope. Keep it consistent anyway: `ES` and `ESP` are two
different scopes, and each would build its own separate history.

### Country registry

Keep this current — it is the only place the mapping is written down.

The five countries in scope, plus the original prototype flow:

| `country` | Teams channel | Document folder |
|---|---|---|
| `BENELUX` | `0_BENELUX DP Comm Channel` | `/Alerts - BENELUX` |
| `FRANCE` | `0_FRANCE DP Comm Channel` | `/Alerts - FRANCE` |
| `GB` | `0_GB DP Comm Channel` | `/Alerts - GB` |
| `IBERIA` | `0_IBERIA DP Comm Channel` | `/Alerts - IBERIA` |
| `IRELAND` | `0_IRELAND DP Comm Channel` | `/Alerts - IRELAND` |
| *(blank — prototype)* | Ai Agents | `/Ai Agents` |

> The original `Weekly Demand Digest` flow over `/Ai Agents/DBAlerts.xlsx` sends no
> `country` and its data showed market `ES40` — i.e. the same business as `IBERIA`.
> Leaving both running produces two digests of the same data in two channels, each
> building its own trend history and each looking correct in isolation. Retire it,
> or give it a distinct scope.
>
> Channel display names and document folder names drift apart — a renamed channel
> keeps its original folder. The **folder** is what appears in the file path, so
> always confirm the folder rather than assuming it matches the channel.
>
> Folders follow `/Alerts - <COUNTRY>`, with spaces around the hyphen, which
> URL-encode as `Alerts%20-%20COUNTRY`. Even so, copy each workbook URL from that
> file's address bar rather than editing another country's — the URL carries a
> per-file id, not just the path.

### Adding a country

> ⚠️ **Build the flow from scratch. Do NOT copy an existing one.**
> A copied flow keeps the original's internal file id. Re-picking the File in the
> designer does not reliably replace it, and the symptom is brutal: `Run script`
> reports **Succeeded**, returns a full execution summary, and never touches the
> workbook you think it targets. The four `List rows` steps then read that
> workbook's stale tables and the digest reports whatever was last left there.
> Rebuilding from scratch fixed it; hours of re-picking did not.

**Build `Run script` alone first, and prove it before adding anything else:**

1. New scheduled flow → add **Run script** only → choose the file by navigating
   Location → Document Library → `/Alerts - <COUNTRY>` → pick `Demand Alerts`.
2. **Save**, then **Test**.
3. Confirm the workbook's **Modified** timestamp moves, and that the note in an
   empty alert sheet shows the current week's Monday.
4. Only now add the Delay, the four `List rows`, the HTTP and the Teams steps.

Ten minutes spent proving step 3 saves a day of debugging a flow whose every
individual setting looks correct.

Then the five things that differ per country:

| # | Where | Change |
|---|---|---|
| 1 | `Run script` | the country's own workbook |
| 2 | 4 × `List rows` | the same workbook (table names are identical everywhere) |
| 3 | HTTP body | `"country":"PT"` |
| 4 | HTTP body | `"workbook_url":"…the PT workbook…"` |
| 5 | Teams step | the country's own channel |

Everything else — URI, token, table names, the Delay, pagination — is identical
across countries.

The Office Script needs no per-country change: each flow references the **same**
script by id and runs it against whichever workbook that flow points at.

> ⚠️ **Thresholds are therefore GLOBAL, not per country.** One script serves every
> country, so editing `ALERT1_MIN_ABS_DIFF` changes all of them at once. A
> materiality that suits IBERIA may swamp or silence a smaller market. If a country
> needs its own thresholds, save a separate copy of the script and point that
> country's `Run script` action at the copy — then remember there are two scripts
> to maintain.

### Verify after adding one

```sql
select country, week_of, count(*) rows, max(created_at) last_write
from demand_alert_snapshots
group by country, week_of
order by week_of desc, country;
```

Each country must appear as its **own row**. If a country is missing, its flow is
not sending `country`; if two countries share a row, they are sending the same
label.

Then confirm each digest survived:

```sql
select country, week_of, model, left(narrative, 60) from demand_digests
order by week_of desc, country;
```

One row per country per week. Fewer rows than countries means the overwrite bug —
migration 002 has not been applied.

Per-country retrieval:

```powershell
Invoke-RestMethod -Uri "https://demand-alerts.onrender.com/api/demand-digest/latest?country=PT" `
  -Headers @{ "X-Digest-Token" = $t }
```

> ⚠️ **The first week of a new country reports no recurring items**, correctly —
> it has no history of its own. It does *not* inherit the trend history of
> countries already running.

---

## Operating it

### Weekly

Nothing. Glance at the Teams post.

### When a narrative looks wrong

Re-send with `"include_brief": true` in the HTTP body. The response then contains
the exact brief the model was given.

- **Brief right, narrative wrong** → the model misread it. Tighten `PROMPT` in
  `digest-bot/digest/digest.py`.
- **Brief wrong** → the bug is upstream, in `digest-bot/digest/alerts.py` or the
  Office Script. Trace the figure back to the worksheet.

The model is instructed to use only figures from the brief, but instructions are not
guarantees. The spot-check in Phase 7 is what catches a drifting model.

### Troubleshooting

| Symptom | Cause |
|---|---|
| `401 Unauthorized` | `X-Digest-Token` missing or not matching `DEMAND_DIGEST_TOKEN` on Render. |
| `No table was found with the name '{GUID}'` | Table field holds an ID, not a name. See 6.4. |
| `Run script` Succeeded, next step "no table found" | Connector cache. Add / lengthen the Delay. See 6.3. |
| `GetTable … NotFound` / `ItemNotFound` when **saving** | Tables don't exist yet, or the File reference is stale. See 6.0. |
| Exactly 256 rows per alert | Pagination off. See 6.4. |
| `metric` / `volume` null in snapshots | Crossed `alertN` mapping. See 6.5. |
| An alert always reports zero rows | That `alertN` may be hardcoded `[]`, or the threshold is too high. |
| Literal `**` in the Teams message | Step 9 uses `narrative` instead of `narrative_html`. |
| Literal `body('HTTP')?['narrative_html']` in Teams | Typed as text, not entered as an expression. See 6.6. |
| Nothing ever reported as recurring | Migration not applied, or publishable Supabase key. Check the Render log for `snapshot/trend step failed`. |
| `column demand_digests.country does not exist` | Migration `002_country.sql` not applied. Digests still post but nothing persists. |
| Only one country's digest per week survives | A flow is not sending `country`, or 002 not applied. See Phase 8. |
| A country's Teams post links to another country's workbook | That flow's `workbook_url` was not changed when the flow was copied. |
| Flow green, `Run script` succeeded, but the workbook is **untouched** | The action's item id points at a different file. The `metadata` path is a cached label and can disagree with `parameters.file`. Confirm by deleting one alert worksheet and re-running: if it does not come back, the flow writes elsewhere. Find the real target by searching SharePoint for the file name and checking Modified. |
| A new country claims items are "recurring" in week 1 | Two flows are sending the **same** `country` label. |
| Digest describes last week's data | Flow ran before the IBP extract refreshed. |
| Worksheet exists with no table | A write failed mid-publish. Re-run; lower `WRITE_CHUNK_CELLS`. See 1.4. |
| Flow times out | Render cold start; raise the HTTP action's timeout. |
| `Digest failed: GROQ_API_KEY is not set` | Env var missing on Render. |
| `413 ... tokens per minute (TPM): Limit 8000, Requested 8480` | Groq bills prompt **plus the whole `max_tokens` reservation** against the per-minute cap. The client now caps the reservation to fit, but a very large brief leaves little room — lower `DEMAND_TOP_N` (try `5`). `GROQ_TPM_LIMIT` raises the assumed cap if the account is upgraded. |
| Teams post stops mid-sentence, later alerts missing | The narrative hit the model's output-token limit. Raise `DEMAND_MAX_OUTPUT_TOKENS`, or lower `DEMAND_TOP_N` to shorten the brief. The service now retries once with a bigger budget and logs `WARNING: the narrative was still truncated` if it remains cut. |
| `model_not_found` / `does not exist or you do not have access` | Groq retired the model id. List current ones with `GET https://api.groq.com/openai/v1/models` (Bearer your key) and update `GROQ_CHAT_MODEL`. The Llama 3.x line disappeared in Aug 2026. |
| HTTP step runs for minutes without finishing | Render free-tier cold start (~30 s) plus snapshot writes on a large payload. Raise the HTTP action's timeout; watch the Render logs to confirm it is progressing rather than stuck. |

### Tuning

| Want | Change |
|---|---|
| Fewer / more alerts | Threshold constants in the **Office Script**. The digest narrates whatever the alerts contain. |
| More items named in the narrative | `DEMAND_TOP_N` on Render. |
| Longer trend memory | `DEMAND_TREND_WEEKS` on Render. |
| Different tone or sections | `PROMPT` in `digest-bot/digest/digest.py`. |
| Workbook link text or target | `WORKBOOK_URL`, or `append_workbook_link` in `digest.py`. |

After changing anything in `digest-bot`, run `tests\test_digest.py` and push —
Render redeploys on push to `main`.

---

## Rebuild checklist

Condensed, for when you already know the build.

```text
[ ] 1  Paste DemandAlerts.ts -> Automate, name "Demand Alerts", Save, Run
[ ]    Read the 7 Dimension lines; confirm "All four output tables verified present."
[ ]    Set ALERT1_MIN_ABS_DIFF / ALERT3_MIN_* to real materiality
[ ]    Paste ListTables.ts, run, confirm 4/4 FOUND
[ ] 2  Run sql/001_demand_alerts.sql then sql/002_country.sql; verify 2 tables + RPC
[ ]    Copy the project URL + the SECRET key
[ ] 3  venv, .env, tests\test_digest.py -> "all checks passed"
[ ] 4  git init/commit; empty GitHub repo; remote set-url; verify remote -v; push
[ ] 5  Render Web Service, Root Directory = digest-bot, 9 env vars, Live, /health ok
[ ]    Copy the REAL service URL from the dashboard
[ ] 6  Flow: Recurrence -> Run script -> Delay 1 min -> 4x List rows -> HTTP -> Teams
[ ]    Rename each List rows after its table
[ ]    Table field = typed NAME (Peek code to confirm, not a GUID)
[ ]    Pagination ON, threshold 5000, on all four
[ ]    HTTP body: alertN -> the action reading tblAlertN
[ ]    Teams message via the fx Expression tab, narrative_html
[ ] 7  Verify: not 256 rows, snapshots stored, 3 figures traced to the worksheets
[ ]    SQL check: metric/volume not null

Per extra country (Phase 8) — copy the flow, change five things:
[ ]    Run script      -> that country's workbook
[ ]    4x List rows    -> the same workbook (table names never change)
[ ]    HTTP body       -> "country":"XX"
[ ]    HTTP body       -> "workbook_url":"...that country's file..."
[ ]    Teams step      -> that country's channel
[ ]    Verify: one snapshot row group and one digest row PER COUNTRY per week
```
