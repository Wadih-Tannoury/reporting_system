"""
Canceled shipments extract - tenant 30018
=========================================

Kibo only, no BigQuery. Walks every site of the tenant, pages through the
shipments updated inside the date window, and reports two things:

  FULL     shipmentStatus == CANCELED
  PARTIAL  any other status, but canceledItems is not empty -- some lines of
           the shipment were canceled while the rest went on

One Excel row per canceled item.

Location history
----------------
A canceled line has usually been passed around before dying: store A rejects
it, it is reassigned to a child shipment, that child cancels it. The child
often carries no fulfillmentLocationCode at all, so the cancellation row on
its own says nothing about where the item actually sat.

item_location_history therefore walks every shipment of the same order that
carried the same product, in chronological order, and lists the fulfillment
location of each. For the DJ-EU-00367393 example that is the store that held
the item (DJITST0002) followed by the unassigned shipment that canceled it.

Date window
-----------
The default window is the last LOOKBACK_HOURS hours (24) ending now, UTC, so a
daily run reports the cancellations of the past day. It can be overridden:

    python kibo_cancellation_report.py --hours 48
    python kibo_cancellation_report.py --from 2026-09-01 --to 2026-10-14

Dates accept YYYY-MM-DD or a full ISO timestamp and are treated as UTC.

Credentials
-----------
Read from the KIBO_API_CREDS environment variable, which holds the JSON
{"client_id": "...", "client_secret": "..."} and may optionally carry "host"
and "auth_url". In GitHub Actions it comes from the repository secret of the
same name. A --credentials-file path still works for local runs.

The window the API is asked for is the requested one widened by
UPDATE_WINDOW_PADDING_DAYS on each side, because Kibo can only filter on the
shipment's last update, which drifts away from the cancellation date whenever
the shipment is touched afterwards. Rows are then filtered precisely on their
own cancellation date.

Columns
-------
    cancellation_type        FULL | PARTIAL
    order_number             externalOrderId
    shipment_order_number    data.shipmentOrderNumber
    order_submission_date    orderSubmitDate
    cancellation_date        canceledItems[].auditInfo.updateDate
    cancellation_reason      shipmentStatusReason / canceledReason reasonCode
    more_info                shipmentStatusReason / canceledReason moreInfo
    product_code             canceledItems[].productCode
    quantity                 canceledItems[].quantity
    line_item_cost           lineItemCost for that product
    fulfillment_location_code      of the shipment that canceled the item
    item_location_history          every location the item passed through
    item_location_count            how many of them
    is_full_order_cancelled  Yes when nothing is left to ship, else No
    shipment_status
    site_id, shipment_number

Requirements
------------
    pip install pandas openpyxl requests
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd
import requests

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name, "").strip()
    return int(raw) if raw else default


def _env_bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "y", "on")


# The secret holding {"client_id": ..., "client_secret": ...}.
KIBO_CREDENTIALS_ENV = "KIBO_API_CREDS"

TENANT_ID = _env_int("KIBO_TENANT_ID", 30018)

DEFAULT_SITE_IDS: Sequence[int] = (
    53592,
    53593,
    53594,
    53595,
    53596,
    53597,
    53598,
    53599,
    53600,
    53601,
)

SITE_IDS: Sequence[int] = tuple(
    int(part) for part in os.environ.get("KIBO_SITE_IDS", "").replace(",", " ").split()
) or DEFAULT_SITE_IDS

KIBO_HOST = "tp3.mozu.com"  # overridden by a "host" key in the credentials file
KIBO_AUTH_URL = "https://home.mozu.com/api/platform/applications/authtickets/oauth"

# --- date window ----------------------------------------------------------- #
# Rolling window: the LOOKBACK_HOURS hours ending now (UTC). --from / --to
# override it for a backfill.
LOOKBACK_HOURS = _env_int("LOOKBACK_HOURS", 24)
# The server-side window on auditInfo.updateDate is widened by this many days
# on each side; rows are then filtered on their real cancellation date. A
# shipment canceled yesterday is often touched again afterwards, which pushes
# its updateDate past the window, so the padding is what keeps it visible.
UPDATE_WINDOW_PADDING_DAYS = _env_int("UPDATE_WINDOW_PADDING_DAYS", 7)
# Backwards padding is nearly pointless: an update cannot predate the
# cancellation it records. One day covers clock skew.
UPDATE_WINDOW_PADDING_BACK_DAYS = _env_int("UPDATE_WINDOW_PADDING_BACK_DAYS", 1)
# Carry the queried window's end up to now, so a cancellation inside the window
# is still found when its shipment was touched again long afterwards. Free for
# the daily run; it is what makes an old backfill complete. Set false to cap the
# end at UPDATE_WINDOW_PADDING_DAYS and keep a far-past backfill short.
PAD_FORWARD_TO_NOW = _env_bool("PAD_FORWARD_TO_NOW", True)

# Discard rows whose own cancellation_date falls outside the requested window.
FILTER_ROWS_BY_CANCELLATION_DATE = True

DATE_FILTER_FIELD: Optional[str] = "auditInfo.updateDate"  # None -> no date filter

# Kibo parses the filter with RSQL/FIQL, which accepts the symbolic operators
# (>=, <=, ==, !=) or the FIQL forms (=ge=, =le=). The word forms "ge"/"le"
# are rejected with a ParseException.
DATE_OP_FROM = ">="
DATE_OP_TO = "<="

# Any further server-side filter, in Kibo filter syntax, appended to the above.
# The status itself is no longer filtered server-side: a partial cancellation
# leaves the shipment in READY, FULFILLED, etc., so every shipment in the
# window has to be inspected for a non-empty canceledItems array.
EXTRA_FILTER: Optional[str] = None

# Report shipments whose status is CANCELED (full cancellations) / whose status
# is something else but that carry canceled items (partial cancellations).
INCLUDE_FULL = True
INCLUDE_PARTIAL = True

# --- channel exclusion ----------------------------------------------------- #
# Kibo carries the channel per item, at data.channel ("ecommerce", "B2B", ...),
# not on the shipment. A row whose channel matches one of these, case
# insensitively, is dropped. An item with no channel of its own inherits the
# shipment's; a row whose channel cannot be determined at all is kept, so an
# unexpected payload shape never silently empties the report.
EXCLUDED_CHANNELS = {
    part.strip().lower()
    for part in os.environ.get("EXCLUDE_CHANNELS", "B2B").split(",")
    if part.strip()
}
CHANNEL_FIELD = "channel"

# --- location history ------------------------------------------------------ #
# Ask Kibo for the other shipments of each affected order, so the history is
# complete even when a sibling shipment fell outside the date window. Set to
# False to build the history only from what the walk already returned, which
# costs no extra calls but can be truncated.
FETCH_ORDER_HISTORY = True
ORDER_LOOKUP_FIELD = "externalOrderId"
HISTORY_SEPARATOR = " > "
# Label used for a shipment that carried the item but has no location code.
UNASSIGNED_LABEL = "(unassigned)"
# Drop those placeholder entries from the history instead of showing them.
HIDE_UNASSIGNED_IN_HISTORY = False
# Collapse a location repeated back to back (A > A > B becomes A > B).
COLLAPSE_REPEATED_LOCATIONS = True

# One row per (shipment_number, product_code). Set to False to keep every
# canceled item line as returned, duplicates included.
DEDUPLICATE = True

# --- paging ---------------------------------------------------------------- #
# Offset paging only works against a deterministic order. Without a sortBy the
# API hands back an arbitrary slice per request, so page 2 overlaps page 1 and
# most of the result set is never reachable. Set to None only to test what the
# unsorted endpoint does.
SORT_BY: Optional[str] = "shipmentNumber asc"
# Page by "shipmentNumber gt <last seen>" rather than by startIndex. Offset
# paging on this API repeats earlier pages past the first page or two, which
# silently loses rows. Requires SORT_BY to order by the same field.
KEYSET_PAGING = _env_bool("KEYSET_PAGING", True)
# Kibo's filter parser is RSQL and wants symbolic comparisons, the same ones
# the date clauses use. The word forms ("gt") are rejected with HTTP 400.
KEYSET_OP = os.environ.get("KEYSET_OP", ">").strip() or ">"

# The window is walked in slices of this many days, so no single query has to
# be paged very deep. A site with thousands of shipments in the window is the
# case this protects. Set to 0 to query the whole window at once.
CHUNK_DAYS = _env_int("CHUNK_DAYS", 3)

# Keep paging past a page that brought nothing new, up to this many times in a
# row, before declaring the slice exhausted.
MAX_EMPTY_PAGES = 2

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "output"))
OUTPUT_PATH = OUTPUT_DIR / (
    f"kibo_canceled_shipments_{TENANT_ID}_"
    f"{datetime.now(timezone.utc):%Y%m%d_%H%M%S}.xlsx"
)

# Write the workbook even when the window is empty, so a scheduled run always
# produces an artifact with the expected three sheets instead of nothing.
ALWAYS_WRITE_FILE = _env_bool("ALWAYS_WRITE_FILE", True)

API_TIMEOUT = 60
API_MAX_RETRIES = 3
API_RETRY_BACKOFF = 2.0
API_PAGE_SIZE = 200
MAX_PAGES_PER_SITE = 2000  # safety stop; the walk is wider than it used to be
# A slice that comes back short is halved and retried. Deep offset paging on
# this API repeats earlier pages instead of returning new rows, so narrowing
# the date range is the only way to reach the remainder. Depth 12 turns a
# 3-day slice into roughly one-minute windows.
MAX_SLICE_SPLIT_DEPTH = _env_int("MAX_SLICE_SPLIT_DEPTH", 48)
MIN_SLICE_SPAN = timedelta(seconds=_env_int("MIN_SLICE_SPAN_SECONDS", 2))
# Bounds for the shipment-number bisection. The ceiling only has to exceed any
# real shipment number; the search narrows to the populated range in a few
# steps and costs nothing when a slice comes back whole.
NUMBER_RANGE_FLOOR = _env_int("NUMBER_RANGE_FLOOR", 0)
NUMBER_RANGE_CEILING = _env_int("NUMBER_RANGE_CEILING", 2_000_000_000)
ORDER_HISTORY_PAGE_SIZE = 100

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("kibo_canceled_shipments")


# --------------------------------------------------------------------------- #
# Dates
# --------------------------------------------------------------------------- #

def _as_utc(value: datetime) -> datetime:
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc)


def parse_cli_date(text: str) -> datetime:
    """YYYY-MM-DD or a full ISO-8601 timestamp, read as UTC."""
    cleaned = text.strip().replace("Z", "+00:00")
    try:
        return _as_utc(datetime.fromisoformat(cleaned))
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Not a date I can read: {text!r}. Use YYYY-MM-DD or an ISO timestamp."
        ) from exc


DATE_ONLY_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def parse_cli_end_date(text: str) -> datetime:
    """
    The window end, where a bare YYYY-MM-DD means the end of that day.

    "--to 2026-05-08" reads naturally as "through the 8th", but a bare date
    parses to midnight, which would silently drop every cancellation made that
    day. A full timestamp is taken exactly as written.
    """
    value = parse_cli_date(text)
    if DATE_ONLY_RE.match(text.strip()):
        value = value.replace(hour=23, minute=59, second=59)
    return value


def _parse_timestamp(value: Any) -> Optional[datetime]:
    if isinstance(value, datetime):
        return _as_utc(value)
    if isinstance(value, str) and value:
        try:
            return _as_utc(datetime.fromisoformat(value.replace("Z", "+00:00")))
        except ValueError:
            return None
    return None


def _kibo_timestamp(value: datetime) -> str:
    """Kibo expects an ISO-8601 UTC timestamp, e.g. 2026-09-14T00:00:00Z."""
    return _as_utc(value).strftime("%Y-%m-%dT%H:%M:%SZ")


def resolve_window(args: argparse.Namespace) -> Tuple[datetime, datetime]:
    """
    Explicit --from / --to win. Otherwise the window is the last `hours` hours
    ending now, which is what the scheduled run uses.
    """
    hours = args.hours if args.hours is not None else LOOKBACK_HOURS
    now = datetime.now(timezone.utc).replace(microsecond=0)

    end = _as_utc(args.date_to) if args.date_to else now
    start = _as_utc(args.date_from) if args.date_from else end - timedelta(hours=hours)
    if end <= start:
        raise SystemExit(
            f"Empty window: --to ({end:%Y-%m-%d %H:%M}) is not after --from ({start:%Y-%m-%d %H:%M})."
        )
    return start, end


def build_filter(start: datetime, end: datetime) -> str:
    """The updateDate window for one slice, plus any extra clause."""
    clauses: List[str] = []

    if DATE_FILTER_FIELD:
        clauses.append(f"{DATE_FILTER_FIELD} {DATE_OP_FROM} {_kibo_timestamp(start)}")
        clauses.append(f"{DATE_FILTER_FIELD} {DATE_OP_TO} {_kibo_timestamp(end)}")

    if EXTRA_FILTER:
        clauses.append(EXTRA_FILTER)

    if not clauses:
        raise SystemExit(
            "No filter would be sent: with DATE_FILTER_FIELD = None and no EXTRA_FILTER "
            "the script would pull every shipment of every site."
        )

    return " and ".join(clauses)


def padded_window(start: datetime, end: datetime) -> Tuple[datetime, datetime]:
    """
    The window actually asked of Kibo, which can only filter on the shipment's
    last update.

    The padding is asymmetric on purpose. A shipment's updateDate can never be
    *earlier* than the cancellation it records, so padding backwards buys
    almost nothing and only costs query time; one day absorbs clock skew.
    Forwards is the direction that matters: a shipment canceled inside the
    window and touched again afterwards carries an updateDate past the window
    end, and a fixed few days of padding silently loses it. The end is
    therefore carried to now, which is free for the daily run (its window
    already ends now) and is what makes a historical backfill complete.
    """
    back = timedelta(days=UPDATE_WINDOW_PADDING_BACK_DAYS)
    forward_end = end + timedelta(days=UPDATE_WINDOW_PADDING_DAYS)
    if PAD_FORWARD_TO_NOW:
        forward_end = max(forward_end, datetime.now(timezone.utc))
    return start - back, forward_end


def window_slices(start: datetime, end: datetime) -> List[Tuple[datetime, datetime]]:
    """
    The padded window cut into CHUNK_DAYS pieces.

    Each piece is queried separately, which keeps every query shallow enough to
    page through reliably. The pieces are half-open apart from a one-second
    step, so a shipment cannot land in two of them.
    """
    first, last = padded_window(start, end)
    if not CHUNK_DAYS or not DATE_FILTER_FIELD:
        return [(first, last)]

    step = timedelta(days=CHUNK_DAYS)
    slices: List[Tuple[datetime, datetime]] = []
    cursor = first
    while cursor < last:
        stop = min(cursor + step, last)
        slices.append((cursor, stop))
        cursor = stop + timedelta(seconds=1)
    return slices or [(first, last)]


# --------------------------------------------------------------------------- #
# Credentials
# --------------------------------------------------------------------------- #

def _first_key(mapping: Dict[str, Any], *candidates: str) -> Optional[Any]:
    lowered = {str(k).lower(): v for k, v in mapping.items()}
    for candidate in candidates:
        value = lowered.get(candidate.lower())
        if value not in (None, ""):
            return value
    return None


def _normalise_credentials(raw: Any, source: str) -> Dict[str, Any]:
    """Expected content: {"client_id": "...", "client_secret": "..."}"""
    if isinstance(raw, dict) and len(raw) == 1:
        only_value = next(iter(raw.values()))
        if isinstance(only_value, dict):
            raw = {**raw, **only_value}

    client_id = _first_key(raw, "client_id", "clientId", "applicationId", "appId")
    client_secret = _first_key(raw, "client_secret", "clientSecret", "sharedSecret", "secret")
    if not client_id or not client_secret:
        raise ValueError(
            f"Could not find client_id / client_secret in {source}. "
            f"Keys present: {sorted(raw) if isinstance(raw, dict) else type(raw).__name__}"
        )

    return {
        "client_id": str(client_id),
        "client_secret": str(client_secret),
        "auth_url": str(_first_key(raw, "authUrl", "auth_url") or KIBO_AUTH_URL),
        "host": str(_first_key(raw, "host", "apiHost", "api_host") or KIBO_HOST),
    }


def load_kibo_credentials_from_env(name: str = KIBO_CREDENTIALS_ENV) -> Dict[str, Any]:
    """The secret, as JSON, straight out of the environment."""
    blob = os.environ.get(name, "").strip()
    if not blob:
        raise SystemExit(
            f"{name} is empty or missing. In GitHub Actions, map the repository "
            f"secret onto the step: env: {name}: ${{{{ secrets.{name} }}}}"
        )
    try:
        raw = json.loads(blob)
    except json.JSONDecodeError as exc:
        # Never echo the value itself.
        raise SystemExit(f"{name} is not valid JSON ({exc.msg} at position {exc.pos}).") from exc

    return _normalise_credentials(raw, f"the {name} secret")


def load_kibo_credentials(path: Path) -> Dict[str, Any]:
    """Same JSON, read from a file instead. Handy for local runs."""
    with path.open("r", encoding="utf-8-sig") as handle:
        raw = json.load(handle)
    return _normalise_credentials(raw, str(path))


# --------------------------------------------------------------------------- #
# Kibo client
# --------------------------------------------------------------------------- #

class KiboClient:
    def __init__(self, credentials: Dict[str, Any], tenant_id: int, filter_expression: str):
        self.client_id = credentials["client_id"]
        self.client_secret = credentials["client_secret"]
        self.auth_url = credentials["auth_url"]
        self.host = credentials["host"]
        self.tenant_id = tenant_id
        self.filter_expression = filter_expression
        # None until the cursor clause has been tried; False if rejected.
        self.keyset_ok: Optional[bool] = None
        self.session = requests.Session()
        self._access_token: Optional[str] = None
        self._token_expires_at: float = 0.0

    # -- auth ---------------------------------------------------------------
    def _authenticate(self) -> None:
        payloads = [
            {
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "grant_type": "client_credentials",
            },
            {"applicationId": self.client_id, "sharedSecret": self.client_secret},
        ]

        data: Dict[str, Any] = {}
        last_error = ""
        for payload in payloads:
            response = self.session.post(self.auth_url, json=payload, timeout=API_TIMEOUT)
            if response.ok:
                data = response.json()
                if data.get("accessToken") or data.get("access_token"):
                    break
            last_error = f"HTTP {response.status_code}: {response.text[:300]}"
            data = {}

        token = data.get("accessToken") or data.get("access_token")
        if not token:
            raise RuntimeError(f"Kibo authentication failed. Last response -> {last_error}")
        self._access_token = token

        expires_in = data.get("expires_in") or data.get("accessTokenExpirationSeconds")
        self._token_expires_at = (
            time.time() + float(expires_in) - 60 if expires_in else time.time() + 15 * 60
        )
        log.info("Kibo authentication successful")

    def _token(self) -> str:
        if not self._access_token or time.time() >= self._token_expires_at:
            self._authenticate()
        return self._access_token  # type: ignore[return-value]

    def _headers(self, site_id: int) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self._token()}",
            "x-vol-tenant": str(self.tenant_id),
            "x-vol-site": str(site_id),
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    # -- shipments ----------------------------------------------------------
    def _request_page(
        self,
        site_id: int,
        start_index: int,
        filter_expression: Optional[str] = None,
        page_size: int = API_PAGE_SIZE,
    ) -> Optional[dict]:
        url = f"https://t{self.tenant_id}.{self.host}/api/commerce/shipments"
        params = {
            "filter": filter_expression if filter_expression is not None else self.filter_expression,
            "pageSize": page_size,
            "startIndex": start_index,
        }
        if SORT_BY:
            params["sortBy"] = SORT_BY

        for attempt in range(1, API_MAX_RETRIES + 1):
            try:
                response = self.session.get(
                    url, headers=self._headers(site_id), params=params, timeout=API_TIMEOUT
                )
            except requests.RequestException as exc:
                log.warning("  site %s @%s attempt %s: %s", site_id, start_index, attempt, exc)
                time.sleep(API_RETRY_BACKOFF * attempt)
                continue

            if response.status_code == 401:
                self._authenticate()
                continue
            if response.status_code in (429, 500, 502, 503, 504):
                log.warning(
                    "  site %s @%s attempt %s: HTTP %s",
                    site_id,
                    start_index,
                    attempt,
                    response.status_code,
                )
                time.sleep(API_RETRY_BACKOFF * attempt)
                continue
            if response.status_code == 404:
                return None
            if not response.ok:
                log.warning(
                    "  site %s @%s: HTTP %s %s",
                    site_id,
                    start_index,
                    response.status_code,
                    response.text[:300],
                )
                return None

            return response.json()

        log.error("  site %s @%s: giving up after %s attempts", site_id, start_index, API_MAX_RETRIES)
        return None

    @staticmethod
    def _shipments_of(payload: Optional[dict]) -> List[dict]:
        if not payload:
            return []
        return (payload.get("_embedded") or {}).get("shipments") or []

    def _page_through(
        self, site_id: int, filter_expression: str, seen_numbers: set, label: str
    ) -> Tuple[List[dict], Optional[int], bool, int]:
        """
        One filter, paged to exhaustion.

        Paging is by shipment number, not by offset. The results are sorted by
        shipmentNumber ascending, so each page asks for the rows after the
        highest number already seen. That is immune to the drift that breaks
        offset paging here, and unlike narrowing the date range it still works
        when hundreds of shipments carry the identical updateDate, as they do
        on bulk-update days.

        Returns the fresh shipments, the total the API claimed, whether the
        walk ended early, and how many distinct rows the slice showed.
        """
        use_keyset = KEYSET_PAGING and self.keyset_ok is not False
        collected: List[dict] = []
        slice_seen: set = set()   # every number this slice returned, new or not
        start_index = 0
        cursor: Optional[int] = None
        total: Optional[int] = None
        empty_streak = 0
        truncated = False

        for _ in range(MAX_PAGES_PER_SITE):
            expression = filter_expression
            if use_keyset and cursor is not None:
                expression = f"{filter_expression} and shipmentNumber {KEYSET_OP} {cursor}"

            payload = self._request_page(site_id, start_index, filter_expression=expression)
            if payload is None:
                if use_keyset and cursor is not None and self.keyset_ok is not False:
                    # The cursor clause itself is being rejected. Disable it for
                    # the rest of the run and redo this slice with offsets,
                    # rather than failing every page from here on.
                    self.keyset_ok = False
                    log.warning(
                        "  %s: the server rejected the '%s' cursor clause; "
                        "falling back to offset paging for the whole run",
                        label, KEYSET_OP,
                    )
                    # Rows already gathered in this attempt are in
                    # seen_numbers, so the retry cannot pick them up again.
                    # Carry them over rather than losing them.
                    again, again_total, again_trunc, again_seen = self._page_through(
                        site_id, filter_expression, seen_numbers, label
                    )
                    return (
                        collected + again,
                        again_total if again_total is not None else total,
                        again_trunc,
                        max(len(slice_seen), again_seen),
                    )
                truncated = True
                log.warning("  %s: page at %s failed", label, start_index)
                break

            page = self._shipments_of(payload)
            if not page:
                break

            page_info = payload.get("page") or {}
            # Only the uncursored first page counts the whole slice; later
            # pages report what is left after the cursor.
            if page_info.get("totalElements") is not None and (
                total is None or not use_keyset
            ):
                total = page_info["totalElements"]

            fresh = []
            for shipment in page:
                number = shipment.get("shipmentNumber")
                if number is None:
                    continue
                slice_seen.add(number)
                if number in seen_numbers:
                    continue
                seen_numbers.add(number)
                fresh.append(shipment)
            collected.extend(fresh)

            if fresh:
                empty_streak = 0
            else:
                empty_streak += 1
                if empty_streak > MAX_EMPTY_PAGES:
                    truncated = total is not None and len(collected) < total
                    break

            if total is not None and len(slice_seen) >= total:
                break

            if use_keyset:
                numbers = [s.get("shipmentNumber") for s in page]
                numbers = [n for n in numbers if isinstance(n, int)]
                if not numbers:
                    truncated = True
                    break
                highest = max(numbers)
                if cursor is not None and highest <= cursor:
                    # The cursor would not advance, so the next request would
                    # return this same page forever.
                    truncated = True
                    break
                cursor = highest
                if len(page) < API_PAGE_SIZE:
                    break
                continue

            if total is not None and start_index + len(page) >= total:
                # The offsets say the result set is exhausted, but fewer rows
                # than the server claimed actually arrived: a later page
                # repeated an earlier one. Walking further only repeats it
                # again, so stop and report the shortfall rather than exiting
                # as if the slice were complete.
                truncated = True
                break

            start_index += len(page)
        else:
            truncated = True

        if total is not None and len(slice_seen) < total:
            truncated = True

        return collected, total, truncated, len(slice_seen)

    def iter_shipments(self, site_id: int, slices: Sequence[Tuple[datetime, datetime]]) -> List[dict]:
        """
        Every shipment of one site in the window, one date slice at a time.

        Slicing keeps each query shallow, which matters because offset paging
        over a large result set is fragile even with a sort: a few thousand
        rows in one query is where it starts to drift.
        """
        shipments: List[dict] = []
        seen_numbers: set = set()
        claimed = 0
        incomplete = False

        def walk(
            slice_start: datetime,
            slice_end: datetime,
            label: str,
            lo: Optional[int] = None,
            hi: Optional[int] = None,
            depth: int = 0,
        ) -> Tuple[int, bool]:
            """
            One slice, narrowed by shipment-number range whenever the server
            returns fewer rows than it claims.

            Splitting on the number rather than on the date is what makes this
            sound. A date range cannot separate rows that share an updateDate,
            and a keyset cursor only works if the server really does sort by
            the cursor field. Shipment numbers are distinct integers, so
            halving the number range always partitions the set exactly, no
            matter what order the server returns rows in or how many share a
            timestamp.
            """
            expression = build_filter(slice_start, slice_end)
            if lo is not None:
                expression += f" and shipmentNumber {KEYSET_OP} {lo}"
            if hi is not None:
                expression += f" and shipmentNumber <= {hi}"

            collected, total, truncated, observed = self._page_through(
                site_id, expression, seen_numbers, label
            )
            shipments.extend(collected)

            short = truncated or (total is not None and observed < total)
            if not short:
                return (total or 0), False

            if total is None:
                # The request itself failed, so the server told us nothing
                # about how many rows there are. Narrowing the number range
                # would just repeat the same failure on both halves.
                log.warning(
                    "  %s: the query failed outright; not splitting, since a "
                    "narrower range would fail the same way",
                    label,
                )
                return 0, True

            lo_eff = NUMBER_RANGE_FLOOR if lo is None else lo
            hi_eff = NUMBER_RANGE_CEILING if hi is None else hi
            if depth >= MAX_SLICE_SPLIT_DEPTH or hi_eff - lo_eff <= 1:
                log.warning(
                    "  %s: %s of %s row(s) unreachable, range cannot be split further",
                    label,
                    (total - observed) if total is not None else "some",
                    total,
                )
                return (total or 0), True

            midpoint = (lo_eff + hi_eff) // 2
            log.info(
                "  %s: got %s of %s, splitting shipmentNumber at %s",
                label, observed, total if total is not None else "?", midpoint,
            )
            left_total, left_bad = walk(
                slice_start, slice_end, f"{label} <{depth + 1}a>", lo_eff, midpoint, depth + 1
            )
            right_total, right_bad = walk(
                slice_start, slice_end, f"{label} <{depth + 1}b>", midpoint, hi_eff, depth + 1
            )
            return left_total + right_total, (left_bad or right_bad)

        for index, (slice_start, slice_end) in enumerate(slices, start=1):
            label = f"site {site_id} [{index}/{len(slices)}] {slice_start:%Y-%m-%d}"
            before = len(shipments)
            slice_claimed, bad = walk(slice_start, slice_end, label)
            claimed += slice_claimed
            if bad:
                incomplete = True

            log.info(
                "  %s: +%s shipment(s) (site total %s, slice claimed %s)",
                label,
                len(shipments) - before,
                len(shipments),
                slice_claimed,
            )

        # claimed counts duplicates across slices, so it is an upper bound only.
        if incomplete:
            log.warning("  site %s: at least one slice was truncated -- counts are a floor", site_id)
        log.info(
            "  site %s: %s distinct shipment(s) collected (slices claimed %s in total)",
            site_id,
            len(shipments),
            claimed,
        )
        return shipments

    def shipments_of_order(self, site_id: int, order_id: str) -> List[dict]:
        """Every shipment of one order, however old, regardless of the window."""
        expression = f'{ORDER_LOOKUP_FIELD} == "{order_id}"'
        shipments: List[dict] = []
        start_index = 0

        seen_numbers: set = set()
        for _ in range(20):  # an order with 2000 shipments does not exist
            payload = self._request_page(
                site_id, start_index, filter_expression=expression, page_size=ORDER_HISTORY_PAGE_SIZE
            )
            page = self._shipments_of(payload)
            if not page:
                break
            fresh = [s for s in page if s.get("shipmentNumber") not in seen_numbers]
            if not fresh:
                break
            for shipment in fresh:
                seen_numbers.add(shipment.get("shipmentNumber"))
            shipments.extend(fresh)

            total = (payload.get("page") or {}).get("totalElements")
            if total is not None and len(shipments) >= total:
                break
            start_index += len(page)

        return shipments


# --------------------------------------------------------------------------- #
# Shipment helpers
# --------------------------------------------------------------------------- #

ITEM_BUCKETS = (
    "items",
    "canceledItems",
    "reassignedItems",
    "rejectedItems",
    "transferredItems",
    "substitutedItems",
)


def line_item_costs(shipment: dict) -> Dict[str, Any]:
    """productCode -> lineItemCost, looking through every item collection."""
    costs: Dict[str, Any] = {}
    for bucket in ITEM_BUCKETS:
        for item in shipment.get(bucket) or []:
            code = item.get("productCode")
            cost = item.get("lineItemCost")
            if code and cost is not None and str(code) not in costs:
                costs[str(code)] = cost
    return costs


def fulfillment_location(shipment: dict) -> Optional[str]:
    """Shipment-level code, falling back to the first item that carries one."""
    code = shipment.get("fulfillmentLocationCode") or shipment.get("assignedLocationCode")
    if code:
        return code
    for bucket in ITEM_BUCKETS:
        for item in shipment.get(bucket) or []:
            if item.get("fulfillmentLocationCode"):
                return item["fulfillmentLocationCode"]
    return None


def item_location(shipment: dict, product_code: Optional[str]) -> Optional[str]:
    """
    Where this particular product sat in this shipment.

    Item-level codes win: a reassigned item carries the location it was
    declined by, which is exactly what the history should show.
    """
    if product_code:
        for bucket in ITEM_BUCKETS:
            for item in shipment.get(bucket) or []:
                if str(item.get("productCode")) == str(product_code) and item.get("fulfillmentLocationCode"):
                    return item["fulfillmentLocationCode"]
    return fulfillment_location(shipment)


def shipment_channel(shipment: dict) -> Optional[str]:
    """
    The shipment's sales channel, taken from its items.

    Item-level data.channel is the sales channel ("Blooming", "ecommerce").
    The shipment's own data.channel is a different thing: it names the
    integration that created the shipment ("Order Connector"), so it is only a
    last resort, used when no item carries a channel at all.
    """
    for bucket in ITEM_BUCKETS:
        for item in shipment.get(bucket) or []:
            item_value = (item.get("data") or {}).get(CHANNEL_FIELD)
            if item_value:
                return str(item_value)
    value = (shipment.get("data") or {}).get(CHANNEL_FIELD)
    return str(value) if value else None


def item_channel(shipment: dict, product_code: Optional[str]) -> Optional[str]:
    """This product's own channel, falling back to the shipment's."""
    if product_code:
        for bucket in ITEM_BUCKETS:
            for item in shipment.get(bucket) or []:
                if str(item.get("productCode")) == str(product_code):
                    value = (item.get("data") or {}).get(CHANNEL_FIELD)
                    if value:
                        return str(value)
    return shipment_channel(shipment)


def is_excluded_channel(channel: Optional[str]) -> bool:
    """True only for a known, excluded channel. Unknown channels are kept."""
    return bool(channel) and channel.strip().lower() in EXCLUDED_CHANNELS


def shipment_order_number_of(shipment: dict) -> Optional[str]:
    data = shipment.get("data") or {}
    return data.get("shipmentOrderNumber") or data.get("xpoOrderId")


def shipment_created_at(shipment: dict) -> Optional[datetime]:
    audit = shipment.get("auditInfo") or {}
    created = _parse_timestamp(audit.get("createDate"))
    if created:
        return created
    # Fall back to the earliest item createDate.
    candidates = [
        _parse_timestamp((item.get("auditInfo") or {}).get("createDate"))
        for bucket in ITEM_BUCKETS
        for item in shipment.get(bucket) or []
    ]
    found = [c for c in candidates if c]
    return min(found) if found else None


def shipment_sort_key(shipment: dict) -> Tuple[datetime, int]:
    """Chronological, with the shipment number as tie-breaker."""
    created = shipment_created_at(shipment) or datetime.max.replace(tzinfo=timezone.utc)
    number = shipment.get("shipmentNumber") or 0
    try:
        number = int(number)
    except (TypeError, ValueError):
        number = 0
    return created, number


def carries_product(shipment: dict, product_code: Optional[str]) -> bool:
    if product_code is None:
        return True
    target = str(product_code)
    for bucket in ITEM_BUCKETS:
        for item in shipment.get(bucket) or []:
            if str(item.get("productCode")) == target:
                return True
    return False


def is_full_cancellation(shipment: dict) -> str:
    """Yes when the shipment is canceled outright or has no line left to ship."""
    if str(shipment.get("shipmentStatus", "")).upper() == "CANCELED":
        return "Yes"
    return "No" if (shipment.get("items") or []) else "Yes"


# --------------------------------------------------------------------------- #
# Location history
# --------------------------------------------------------------------------- #

class LocationHistory:
    """
    Every fulfillment location a canceled item passed through, oldest first.

    The shipments of an order are taken from what the walk already returned and,
    when FETCH_ORDER_HISTORY is on, topped up with a per-order lookup so that a
    sibling shipment outside the date window is not missed.
    """

    def __init__(self, kibo: KiboClient, seen: Dict[Tuple[int, str], Dict[int, dict]]):
        self.kibo = kibo
        self.seen = seen
        self._fetched: set = set()
        self.lookups = 0

    def _order_shipments(self, site_id: int, order_id: Optional[str]) -> List[dict]:
        if not order_id:
            return []
        key = (site_id, str(order_id))
        pool = self.seen.setdefault(key, {})

        if FETCH_ORDER_HISTORY and key not in self._fetched:
            self._fetched.add(key)
            self.lookups += 1
            for shipment in self.kibo.shipments_of_order(site_id, str(order_id)):
                number = shipment.get("shipmentNumber")
                if number is not None:
                    pool[number] = shipment

        return sorted(pool.values(), key=shipment_sort_key)

    def for_item(
        self, site_id: int, order_id: Optional[str], product_code: Optional[str]
    ) -> Tuple[str, int]:
        locations: List[str] = []
        for shipment in self._order_shipments(site_id, order_id):
            if not carries_product(shipment, product_code):
                continue
            code = item_location(shipment, product_code)
            if not code:
                if HIDE_UNASSIGNED_IN_HISTORY:
                    continue
                code = UNASSIGNED_LABEL
            if COLLAPSE_REPEATED_LOCATIONS and locations and locations[-1] == code:
                continue
            locations.append(code)

        real = [c for c in locations if c != UNASSIGNED_LABEL]
        return HISTORY_SEPARATOR.join(locations), len(real)


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #

COLUMNS = [
    "cancellation_type",
    "order_number",
    "shipment_order_number",
    "order_submission_date",
    "cancellation_date",
    "cancellation_reason",
    "more_info",
    "product_code",
    "quantity",
    "line_item_cost",
    "fulfillment_location_code",
    "item_location_history",
    "item_location_count",
    "is_full_order_cancelled",
    "shipment_status",
    "channel",
    "site_id",
    "shipment_number",
]


def extract_rows(
    site_id: int,
    shipments: Sequence[dict],
    history: LocationHistory,
) -> Tuple[List[Dict[str, Any]], int]:
    """
    One record per canceled item, plus the number of rows dropped by channel.

    A shipment qualifies when its status is CANCELED (full cancellation) or
    when it carries canceled items despite another status (partial). Rows on an
    excluded channel (B2B by default) are dropped before the location-history
    lookup, so they cost no extra API calls.
    """
    rows: List[Dict[str, Any]] = []
    excluded = 0

    for shipment in shipments:
        status = str(shipment.get("shipmentStatus", "")).upper()
        canceled_items = shipment.get("canceledItems") or []
        full = status == "CANCELED"

        if full and not INCLUDE_FULL:
            continue
        if not full:
            if not canceled_items or not INCLUDE_PARTIAL:
                continue

        order_id = shipment.get("externalOrderId")
        reason = shipment.get("shipmentStatusReason") or {}
        costs = line_item_costs(shipment)
        location = fulfillment_location(shipment)
        full_order = is_full_cancellation(shipment)
        channel = shipment_channel(shipment)

        common = {
            "cancellation_type": "FULL" if full else "PARTIAL",
            "order_number": order_id,
            "shipment_order_number": shipment_order_number_of(shipment),
            "order_submission_date": shipment.get("orderSubmitDate"),
            "fulfillment_location_code": location,
            "is_full_order_cancelled": full_order,
            "shipment_status": shipment.get("shipmentStatus"),
            "channel": channel,
            "site_id": site_id,
            "shipment_number": shipment.get("shipmentNumber"),
        }

        if not canceled_items:
            # CANCELED with an empty canceledItems array: keep the shipment
            # visible instead of letting it drop out silently.
            if is_excluded_channel(channel):
                excluded += 1
                continue
            audit = shipment.get("auditInfo") or {}
            trail, count = history.for_item(site_id, order_id, None)
            rows.append(
                {
                    **common,
                    "cancellation_date": audit.get("updateDate"),
                    "cancellation_reason": reason.get("reasonCode"),
                    "more_info": reason.get("moreInfo"),
                    "product_code": None,
                    "quantity": None,
                    "line_item_cost": None,
                    "item_location_history": trail,
                    "item_location_count": count,
                }
            )
            continue

        for item in canceled_items:
            audit = item.get("auditInfo") or {}
            item_reason = item.get("canceledReason") or {}
            product_code = item.get("productCode")

            # The line's own channel leads, so a mixed shipment loses only its
            # B2B lines. Checked before the history lookup to save the call.
            line_channel = (item.get("data") or {}).get(CHANNEL_FIELD) or channel
            if is_excluded_channel(line_channel):
                excluded += 1
                continue

            cost = item.get("lineItemCost")
            if cost is None and product_code:
                cost = costs.get(str(product_code))

            # On a partial cancellation the shipment-level reason describes the
            # shipment, not the canceled line, so the item's own reason leads.
            if full:
                reason_code = reason.get("reasonCode") or item_reason.get("reasonCode")
                more_info = reason.get("moreInfo") or item_reason.get("moreInfo")
            else:
                reason_code = item_reason.get("reasonCode") or reason.get("reasonCode")
                more_info = item_reason.get("moreInfo") or reason.get("moreInfo")

            trail, count = history.for_item(site_id, order_id, product_code)

            rows.append(
                {
                    **common,
                    "cancellation_date": audit.get("updateDate"),
                    "cancellation_reason": reason_code,
                    "more_info": more_info,
                    "product_code": product_code,
                    "quantity": item.get("quantity"),
                    "line_item_cost": cost,
                    "fulfillment_location_code": item.get("fulfillmentLocationCode") or location,
                    "item_location_history": trail,
                    "item_location_count": count,
                    "channel": line_channel,
                }
            )

    return rows, excluded


# --------------------------------------------------------------------------- #
# Output
# --------------------------------------------------------------------------- #

def _to_naive_utc(value: Any) -> Any:
    """Excel cannot store tz-aware datetimes."""
    if value is None or value == "":
        return None
    if isinstance(value, str):
        try:
            value = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return value
    if isinstance(value, datetime) and value.tzinfo is not None:
        return value.astimezone(timezone.utc).replace(tzinfo=None)
    return value


def _prepare(records: List[Dict[str, Any]], start: datetime, end: datetime) -> pd.DataFrame:
    frame = pd.DataFrame(records, columns=COLUMNS)
    for column in ("order_submission_date", "cancellation_date"):
        frame[column] = frame[column].map(_to_naive_utc)

    if FILTER_ROWS_BY_CANCELLATION_DATE and not frame.empty:
        dates = pd.to_datetime(frame["cancellation_date"], errors="coerce")
        window = dates.isna() | (
            (dates >= pd.Timestamp(_as_utc(start).replace(tzinfo=None)))
            & (dates <= pd.Timestamp(_as_utc(end).replace(tzinfo=None)))
        )
        dropped = int((~window).sum())
        if dropped:
            log.info("Dropped %s row(s) canceled outside the requested window", dropped)
        frame = frame[window]

    if DEDUPLICATE:
        frame = frame.drop_duplicates(subset=["shipment_number", "product_code"])

    return frame.sort_values(["cancellation_date", "shipment_order_number"], na_position="last")


def _autosize(sheet, frame: pd.DataFrame) -> None:
    for index, column in enumerate(frame.columns, start=1):
        values = [len(str(v)) for v in frame[column].head(200)] or [0]
        width = max(len(str(column)), *values)
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = min(width + 2, 40)


def write_excel(records: List[Dict[str, Any]], path: Path, start: datetime, end: datetime) -> int:
    frame = _prepare(records, start, end)
    sheets = {
        "cancellations": frame,
        "full": frame[frame["cancellation_type"] == "FULL"],
        "partial": frame[frame["cancellation_type"] == "PARTIAL"],
    }

    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="YYYY-MM-DD HH:MM:SS") as writer:
        for name, sheet_frame in sheets.items():
            sheet_frame.to_excel(writer, index=False, sheet_name=name)
            _autosize(writer.sheets[name], sheet_frame)

    log.info(
        "Written %s row(s) (%s full, %s partial) to %s",
        len(frame),
        len(sheets["full"]),
        len(sheets["partial"]),
        path,
    )
    return len(frame)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Canceled shipments extract from Kibo.")
    parser.add_argument("--from", dest="date_from", type=parse_cli_date, default=None,
                        help="Window start, inclusive. YYYY-MM-DD or ISO timestamp, UTC.")
    parser.add_argument("--to", dest="date_to", type=parse_cli_end_date, default=None,
                        help="Window end, inclusive. A bare YYYY-MM-DD means the end of that "
                             "day (23:59:59 UTC); a full ISO timestamp is used as written. "
                             "Defaults to now.")
    parser.add_argument("--hours", type=int, default=None,
                        help=f"Look back this many hours from now. Default {LOOKBACK_HOURS}. "
                             "Ignored when --from is given.")
    parser.add_argument("--trace-order", default=None, metavar="ORDER_NUMBER",
                        help="Explain why one order (externalOrderId, e.g. DJ-EU-00361631) is "
                             "or is not in the report for the given window, then exit.")
    parser.add_argument("--credentials-file", type=Path, default=None,
                        help=f"Read the credentials JSON from this file instead of "
                             f"the {KIBO_CREDENTIALS_ENV} environment variable.")
    parser.add_argument("--no-order-history", action="store_true",
                        help="Build item_location_history only from the shipments the walk returned.")
    parser.add_argument("--output", type=Path, default=OUTPUT_PATH, help="Destination .xlsx path.")
    return parser.parse_args(argv)


def index_shipments(
    index: Dict[Tuple[int, str], Dict[int, dict]], site_id: int, shipments: Sequence[dict]
) -> None:
    """Keep every shipment the walk saw, keyed by (site, order), for the history."""
    for shipment in shipments:
        order_id = shipment.get("externalOrderId")
        number = shipment.get("shipmentNumber")
        if order_id and number is not None:
            index.setdefault((site_id, str(order_id)), {})[number] = shipment


def trace_order(kibo: "KiboClient", order_id: str, start: datetime, end: datetime) -> int:
    """
    Explain, shipment by shipment, why one order is or is not in the report.

    Looks the order up directly, bypassing the date-windowed walk, then checks
    it against every stage a row has to survive: fetched by the server-side
    updateDate filter, qualifies as a cancellation, channel not excluded,
    cancellation date inside the requested window.
    """
    padded_start, padded_end = padded_window(start, end)
    log.info("Tracing order %s", order_id)
    log.info("  requested window : %s -> %s", _kibo_timestamp(start), _kibo_timestamp(end))
    log.info("  queried window   : %s -> %s", _kibo_timestamp(padded_start), _kibo_timestamp(padded_end))

    found_any = False
    reported = 0

    for site_id in SITE_IDS:
        shipments = kibo.shipments_of_order(site_id, order_id)
        if not shipments:
            continue
        found_any = True
        log.info("  site %s: %s shipment(s)", site_id, len(shipments))

        for shipment in sorted(shipments, key=shipment_sort_key):
            number = shipment.get("shipmentNumber")
            status = str(shipment.get("shipmentStatus", "")).upper()
            updated = _parse_timestamp((shipment.get("auditInfo") or {}).get("updateDate"))
            canceled_items = shipment.get("canceledItems") or []
            full = status == "CANCELED"

            log.info("    shipment %s [%s] updateDate=%s canceledItems=%s",
                     number, status,
                     _kibo_timestamp(updated) if updated else "none",
                     len(canceled_items))

            if updated and not (padded_start <= updated <= padded_end):
                log.info("      EXCLUDED: the walk never sees it -- its updateDate is outside "
                         "the queried window. Widen the window or the padding.")
                continue
            if not full and not canceled_items:
                log.info("      not a cancellation: status is not CANCELED and nothing was canceled")
                continue

            channel = shipment_channel(shipment)
            rows = canceled_items or [None]
            for item in rows:
                code = (item or {}).get("productCode") if item else None
                line_channel = ((item or {}).get("data") or {}).get(CHANNEL_FIELD) or channel
                when = _parse_timestamp(
                    ((item or {}).get("auditInfo") or (shipment.get("auditInfo") or {})).get("updateDate")
                )
                label = f"      product {code or '(none)'}"

                if is_excluded_channel(line_channel):
                    log.info("%s EXCLUDED: channel %r is in EXCLUDE_CHANNELS", label, line_channel)
                    continue
                if when is None:
                    log.info("%s INCLUDED: no cancellation date, kept", label)
                    reported += 1
                    continue
                if not (start <= when <= end):
                    log.info("%s EXCLUDED: cancellation_date %s is outside the requested "
                             "window", label, _kibo_timestamp(when))
                    continue
                log.info("%s INCLUDED: cancellation_date=%s channel=%r",
                         label, _kibo_timestamp(when), line_channel)
                reported += 1

    if not found_any:
        log.warning("  order %s was not found on any configured site. Check the order "
                    "number and that its site is in SITE_IDS.", order_id)
        return 1

    log.info("  verdict: %s row(s) would appear in the report for this order", reported)
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    global FETCH_ORDER_HISTORY

    args = parse_args(argv)
    if args.no_order_history:
        FETCH_ORDER_HISTORY = False

    start, end = resolve_window(args)
    log.info("Window: %s -> %s (UTC)", _kibo_timestamp(start), _kibo_timestamp(end))

    if args.credentials_file:
        if not args.credentials_file.exists():
            log.error("Credentials file not found: %s", args.credentials_file)
            return 1
        credentials = load_kibo_credentials(args.credentials_file)
    else:
        credentials = load_kibo_credentials_from_env()

    padded_start, padded_end = padded_window(start, end)
    slices = window_slices(start, end)

    kibo = KiboClient(
        credentials,
        TENANT_ID,
        build_filter(padded_start, padded_end),
    )
    if args.trace_order:
        return trace_order(kibo, args.trace_order, start, end)

    log.info(
        "Scanning %s to %s in %s slice(s) of %s day(s), sorted by %s",
        _kibo_timestamp(padded_start),
        _kibo_timestamp(padded_end),
        len(slices),
        CHUNK_DAYS or "all",
        SORT_BY or "nothing (unstable paging)",
    )

    # Pass 1: walk the sites and keep what came back, so the history can reuse it.
    per_site_shipments: List[Tuple[int, List[dict]]] = []
    seen: Dict[Tuple[int, str], Dict[int, dict]] = defaultdict(dict)

    for position, site_id in enumerate(SITE_IDS, start=1):
        log.info("Site [%s/%s] %s", position, len(SITE_IDS), site_id)
        shipments = kibo.iter_shipments(site_id, slices)
        index_shipments(seen, site_id, shipments)
        per_site_shipments.append((site_id, shipments))

    # Pass 2: build the rows, resolving each item's location history.
    history = LocationHistory(kibo, seen)
    records: List[Dict[str, Any]] = []
    per_site: List[Tuple[int, int, int, int]] = []

    excluded_total = 0
    for site_id, shipments in per_site_shipments:
        rows, excluded = extract_rows(site_id, shipments, history)
        excluded_total += excluded
        full = sum(1 for r in rows if r["cancellation_type"] == "FULL")
        per_site.append((site_id, len(shipments), full, len(rows) - full))
        log.info(
            "  site %s -> %s shipment(s) scanned, %s full + %s partial row(s)%s",
            site_id,
            len(shipments),
            full,
            len(rows) - full,
            f", {excluded} excluded by channel" if excluded else "",
        )
        records.extend(rows)

    if EXCLUDED_CHANNELS:
        log.info(
            "Excluded channel(s): %s -- %s row(s) dropped",
            ", ".join(sorted(EXCLUDED_CHANNELS)),
            excluded_total,
        )

    log.info("Summary (site: shipments scanned / full rows / partial rows)")
    for site_id, scanned, full, partial in per_site:
        log.info("  %s: %s / %s / %s", site_id, scanned, full, partial)
    if FETCH_ORDER_HISTORY:
        log.info("Order history lookups: %s", history.lookups)

    if not records and not ALWAYS_WRITE_FILE:
        log.warning("No cancellations found in the window; nothing to write.")
        return 0
    if not records:
        log.warning("No cancellations found in the window; writing an empty workbook.")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    row_count = write_excel(records, args.output, start, end)
    _publish_outputs(args.output, row_count, start, end)
    return 0


def _publish_outputs(path: Path, rows: int, start: datetime, end: datetime) -> None:
    """Hand the result back to the workflow through $GITHUB_OUTPUT."""
    target = os.environ.get("GITHUB_OUTPUT")
    if not target:
        return
    with open(target, "a", encoding="utf-8") as handle:
        handle.write(f"report_path={path}\n")
        handle.write(f"report_name={path.name}\n")
        handle.write(f"row_count={rows}\n")
        handle.write(f"window_start={_kibo_timestamp(start)}\n")
        handle.write(f"window_end={_kibo_timestamp(end)}\n")


if __name__ == "__main__":
    raise SystemExit(main())
