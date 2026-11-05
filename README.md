# Kibo cancellation report

Daily extract of canceled shipments for Kibo tenant 30018, as an Excel file.

## What it does

Every day at 08:00 Europe/Rome, the workflow walks all ten sites of the tenant,
collects the shipments canceled in the **last 24 hours** (by `cancellation_date`,
not by last update), and emails an `.xlsx` with three sheets:

| Sheet | Contents |
|---|---|
| `cancellations` | every canceled item |
| `full` | shipments whose status is `CANCELED` |
| `partial` | shipments in another status carrying canceled items |

Columns are unchanged from the original script, `cancellation_type` through
`shipment_number`, including `item_location_history`.

The workbook goes out as **DJ daily cancellation report** to:

- giulia.silvestro@ladoublej.com
- sabrina.baraldi@ladoublej.com
- sara.mattioli@ladoublej.com
- giulia.piroli@ladoublej.com
- beatrice.cabrini@ladoublej.com

The list lives at the top of `scripts/send_report_email.py`; a repository
variable `EMAIL_RECIPIENTS` (comma separated) overrides it without a code
change. The same file is also attached to the run as an artifact for 90 days,
so a failed send never loses the data.

On a day with no cancellations the email still goes out, with an empty workbook
and a body saying so.

## Setup

1. Copy `scripts/`, `requirements.txt` and `.github/workflows/` into the repo.
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
python scripts/kibo_cancellation_report.py

# or from a file, over a chosen window
python scripts/kibo_cancellation_report.py \
    --credentials-file ~/kibo_credentials.json \
    --from 2026-09-01 --to 2026-10-14

# email it, logging instead of sending
EMAIL_DRY_RUN=true python scripts/send_report_email.py --file output/<name>.xlsx

# send to yourself only
python scripts/send_report_email.py --file output/<name>.xlsx --to you@thelevelgroup.com
```

## Tuning

Set these as `env:` on the report step, or shell variables locally.

| Variable | Default | Purpose |
|---|---|---|
| `LOOKBACK_HOURS` | `24` | Size of the rolling window |
| `UPDATE_WINDOW_PADDING_DAYS` | `7` | Days added each side of the window when querying Kibo, because the API filters on last update, which drifts past the cancellation date. Rows are then filtered on their real `cancellation_date`. Lower it to shorten the run, at the risk of missing a late-touched shipment |
| `CHUNK_DAYS` | `3` | Slice size for paging |
| `KIBO_TENANT_ID` | `30018` | Tenant |
| `KIBO_SITE_IDS` | the ten sites | Comma-separated override |
| `ALWAYS_WRITE_FILE` | `true` | Write an empty workbook when nothing was canceled, so the artifact always exists. Set `false` to skip the upload instead — the upload step then fails on a quiet day, since `if-no-files-found` is `error` |
| `EMAIL_RECIPIENTS` | the five DJ addresses | Comma-separated override |
| `EMAIL_SUBJECT` | `DJ daily cancellation report` | Overridden by `--subject`, which the workflow passes explicitly |
| `EMAIL_DRY_RUN` | `false` | Log the email instead of sending it |

With the defaults, one run queries 15 days of updates across 10 sites in 3-day
slices, plus a per-order lookup for the location history. Expect roughly 10 to
30 minutes; the job times out at 120.
