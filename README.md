# Kibo cancellation report

Daily extract of canceled shipments for Kibo tenant 30018, as an Excel file.

## What it does

Every day at 08:00 Europe/Rome, the workflow walks all ten sites of the tenant,
collects the shipments canceled in the **last 24 hours** (by `cancellation_date`,
not by last update), and emails an `.xlsx` with three sheets:

Shipments on the **B2B channel are excluded**. Kibo carries the sales channel
per item at `data.channel` (`Blooming`, `ecommerce`); the shipment's own
`data.channel` is the integration that created it (`Order Connector`) and is
used only when no item carries a channel. The filter works line by line: a mixed shipment keeps
its non-B2B lines and loses the rest. A row whose channel cannot be determined
is kept, so an unexpected payload never silently empties the report. A row whose
`cancellation_date` cannot be read is dropped, since the window is what the
report is about and an unreadable date cannot be shown to fall inside it. The run log
reports how many rows were dropped.

| Sheet | Contents |
|---|---|
| `cancellations` | every canceled item |
| `full` | shipments whose status is `CANCELED` |
| `partial` | shipments in another status carrying canceled items |

Columns are those of the original script, `cancellation_type` through
`shipment_number`, including `item_location_history`, plus a `channel` column so
the exclusion can be audited.

The workbook goes out as **DJ daily cancellation report** to:

- giulia.silvestro@ladoublej.com
- sabrina.baraldi@ladoublej.com
- sara.mattioli@ladoublej.com
- giulia.piroli@ladoublej.com
- beatrice.cabrini@ladoublej.com

The list lives at the top of `send_report_email.py`; a repository
variable `EMAIL_RECIPIENTS` (comma separated) overrides it without a code
change. The same file is also attached to the run as an artifact for 90 days,
so a failed send never loses the data.

On a day with no cancellations the email still goes out, with no attachment and
the single line:

> No DJ shipments were cancelled in the last 24 hours

Over a custom window the line names the window instead. The artifact is still
uploaded, so the empty workbook remains available from the run.

## Setup

1. The two scripts sit at the repository root, beside `requirements.txt`.
2. Add a repository secret named `KIBO_API_CREDS` under
   *Settings → Secrets and variables → Actions*:

   ```json
   {
       "client_id": "...",
       "client_secret": "..."
   }
   ```

   `host` and `auth_url` keys are accepted too, and default to `tp3.mozu.com`
   and the standard Mozu OAuth endpoint.
3. Add the SMTP settings. These are the **same names the ERP Alerting System
   uses**, so if both live in the same repository they are already there and
   nothing further is needed.

   | Kind | Name | Example |
   |---|---|---|
   | Secret | `SMTP_USERNAME` | `wadih.tannoury@thelevelgroup.com` |
   | Secret | `SMTP_PASSWORD` | mailbox or app password |
   | Variable | `SMTP_HOST` | `smtp.office365.com` |
   | Variable | `SMTP_PORT` | `587` |
   | Variable | `SMTP_STARTTLS` | `true` |
   | Variable | `EMAIL_FROM` | sender address, defaults to `SMTP_USERNAME` |
   | Variable | `SMTP_AUTH_MECHANISM` | `AUTO`, or `LOGIN` / `PLAIN` to force one |
   | Variable | `SMTP_REMOVE_PASSWORD_SPACES` | `true` if the app password was copied with spaces |
4. Run once by hand with mode **smtp-login-test** to confirm the mailbox
   authenticates, then once with **run** and *dry run* ticked.
5. Push to the default branch. A scheduled workflow only runs from there.

## Running it by hand

*Actions → Kibo cancellation report → Run workflow*:

- **mode** — `run` builds and emails; `report-only` builds the artifact without
  emailing, which is what a backfill usually wants; `smtp-login-test` only
  checks that the mailbox authenticates.
- **hours** — look back this many hours instead of 24.
- **from** / **to** — a specific window, `YYYY-MM-DD` or an ISO timestamp, UTC.
  `from` wins over `hours`, and `to` defaults to now. Use these for a backfill.
  A bare date in **to** means the *end* of that day, so `to = 2026-05-08`
  includes cancellations made during the 8th. Pass a full timestamp
  (`2026-05-08T12:00:00`) to cut the window mid-day instead.
- **dry run** — log the email instead of sending it, so the five recipients are
  not disturbed by a test.

## The two cron lines

GitHub cron is UTC and ignores daylight saving, so 08:00 in Milan moves between
06:00 and 07:00 UTC across the year. Both are scheduled, and the first step of
the job exits early on whichever trigger is not 08:00 local. A skipped run shows
as a green run with one step and no artifact.

Two things to know about GitHub's scheduler: runs can be delayed by several
minutes at busy times, and the schedule is suspended after 60 days without
repository activity, which a manual run resets.

## Local runs

```bash
pip install -r requirements.txt

# last 24 hours, credentials from the environment
export KIBO_API_CREDS='{"client_id":"...","client_secret":"..."}'
python kibo_cancellation_report.py

# or from a file, over a chosen window
python kibo_cancellation_report.py \
    --credentials-file ~/kibo_credentials.json \
    --from 2026-09-01 --to 2026-10-14

# email it, logging instead of sending
EMAIL_DRY_RUN=true python send_report_email.py --file output/<name>.xlsx

# send to yourself only
python send_report_email.py --file output/<name>.xlsx --to you@thelevelgroup.com
```

## Paging

Kibo's offset paging is not reliable here: past the first page or two, a
deeper `startIndex` returns an earlier page again instead of new rows. A slice
claiming 267 rows handed back 200 and stopped, and the other 67 were lost
silently.

Paging is therefore by **shipment number**, not by offset. Results are sorted
by `shipmentNumber asc`, and each page asks for rows after the highest number
already seen (`shipmentNumber > N`). Kibo's filter parser is RSQL and wants
the symbolic comparisons the date clauses already use; the word forms (`gt`)
come back as HTTP 400. The cursor cannot repeat itself, so the
walk reaches the end of a slice however large it is, in fewer calls than
offset paging took to fail. If the server ever rejects the cursor clause, the
run says so once and falls back to offset paging rather than failing every
page. `KEYSET_PAGING=false` forces offsets; `KEYSET_OP` changes the operator.

Slice splitting remains as a safety net: if a slice still returns fewer rows
than the server claims, the date range is halved and retried, bounded by
`MAX_SLICE_SPLIT_DEPTH` (12) and `MIN_SLICE_SPAN_SECONDS` (2). With keyset
paging it should rarely trigger. Splitting alone could not fix bulk-update
days, where hundreds of shipments share one `updateDate` and no date range is
narrow enough to separate them; keyset paging handles those.

## Why is an order missing?

Run the report with `--trace-order` and it will say, stage by stage, what
happened to every shipment of that order:

```bash
python kibo_cancellation_report.py --from 2026-05-20 --to 2026-05-20 \
    --trace-order DJ-EU-00361631
```

Each line ends in `INCLUDED` or `EXCLUDED` with the reason: never fetched
(its last update falls outside the queried window), not a cancellation,
channel excluded, or a `cancellation_date` outside the requested window.

## Tuning

Set these as `env:` on the report step, or shell variables locally.

| Variable | Default | Purpose |
|---|---|---|
| `LOOKBACK_HOURS` | `24` | Size of the rolling window |
| `UPDATE_WINDOW_PADDING_DAYS` | `7` | Days added after the window end when querying Kibo |
| `UPDATE_WINDOW_PADDING_BACK_DAYS` | `1` | Days added before the window start. An update cannot predate the cancellation it records, so this only absorbs clock skew |
| `PAD_FORWARD_TO_NOW` | `true` | Carry the queried end up to now, so a cancellation inside the window is still found when its shipment was touched again weeks later. Free for the daily run; set `false` to keep a far-past backfill short, accepting that late-touched shipments go missing |
| `CHUNK_DAYS` | `3` | Slice size for paging |
| `KIBO_TENANT_ID` | `30018` | Tenant |
| `KIBO_SITE_IDS` | the ten sites | Comma-separated override |
| `ALWAYS_WRITE_FILE` | `true` | Write an empty workbook when nothing was canceled, so the artifact always exists. Set `false` to skip the upload instead — the upload step then fails on a quiet day, since `if-no-files-found` is `error` |
| `KEEP_UNDATED_ROWS` | `false` | Keep rows whose `cancellation_date` cannot be parsed. They are dropped by default, so the report holds only cancellations inside the window |
| `EXCLUDE_CHANNELS` | `B2B` | Comma-separated channels to drop, case insensitive. Set to empty to keep every channel |
| `EMAIL_RECIPIENTS` | the five DJ addresses | Comma-separated override |
| `EMAIL_SUBJECT` | `DJ daily cancellation report` | Overridden by `--subject`, which the workflow passes explicitly |
| `EMAIL_DRY_RUN` | `false` | Log the email instead of sending it |

With the defaults, one run queries 15 days of updates across 10 sites in 3-day
slices, plus a per-order lookup for the location history. Expect roughly 10 to
30 minutes; the job times out at 120.
