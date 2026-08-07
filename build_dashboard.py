#!/usr/bin/env python3
"""
Pathfinder Lead Dashboard — daily data refresh.

Downloads the two live Google Sheets workbooks (both shared "Anyone with the
link"), recomputes lead-contact metrics using the exact-match methodology
agreed with the team, and regenerates index.html from template.html.

Runs on a GitHub Actions daily schedule — no Claude involvement, no manual
steps, no cost beyond GitHub's free Actions minutes.

METHODOLOGY (do not change without re-checking with the team — this took a
few rounds to nail down correctly):
  - "Contacted" = the contact-status column reads EXACTLY "Yes" (phone
    conversation) OR EXACTLY "Couldn't Connect, Sent Whatsapp" (WhatsApp-only
    reply). No other value counts — not blanks, not "No", not near-miss text
    like "No Call, Sent Whatsapp". This is intentional exact-string matching,
    not fuzzy/contains matching.
  - not_contacted = total - contacted
  - "total" = number of rows with a non-empty Name cell (trailing template
    rows in the sheet have no name and are excluded).
  - This rule is applied uniformly across all 5 sheets.
  - Follow-up is a THREE-way classification on the follow-up column (revised
    2026 — this is intentionally NOT a simple Yes/No anymore):
      * Cell reads EXACTLY "Yes"  -> follow-up attempted AND successful
        (a follow-up call was made and a conversation took place).
      * Cell reads EXACTLY "No"   -> follow-up attempted but NOT successful
        (a follow-up attempt was made — e.g. a call — but no conversation
        resulted; the salesperson may have then sent a WhatsApp message or
        taken some other fallback action). This is NOT "no follow-up".
      * Blank cell                -> no follow-up attempt was made at all.
    follow_successful     = count of exactly "Yes"
    follow_unsuccessful   = count of exactly "No"
    follow_attempted      = follow_successful + follow_unsuccessful
    follow_not_attempted  = max(0, contacted - follow_attempted)   -- see note below
    follow_attempt_rate            = follow_attempted / total
    follow_success_rate            = follow_successful / total
    follow_success_rate_of_attempts = follow_successful / follow_attempted
    IMPORTANT: follow_not_attempted is scoped to CONTACTED leads only (revised
    2026-08) — you can't follow up on someone who was never reached, so leads
    that were never contacted are excluded entirely rather than being counted
    as "not followed up". This means successful + unsuccessful + not_attempted
    reconciles back to "contacted", NOT "total" — the gap between "contacted"
    and "total" (i.e. never-contacted leads) sits outside this classification
    altogether. Clamped at 0 as a safety net since "contacted" and "followed
    up" come from two independently-maintained columns.
    "followed" is kept as a field name for backward compatibility elsewhere
    in this pipeline and is now an alias for follow_successful.
  - "Duplicate rows" = each sheet's final column is a manually-maintained
    uniqueness marker that reads exactly "Unique" for every normal row. Any
    row where that cell is non-blank and NOT "Unique" is a flagged duplicate;
    we surface the name/phone/sheet and the flag text as-is.
  - Recent-activity detail rows also carry: "method" (how the lead was
    reached — "Phone call" if the contact column reads exactly "Yes",
    "WhatsApp" if it reads exactly the WhatsApp-only string, else blank) and
    "days_to_follow" (integer days between Date Called and Date Followed Up,
    only set when BOTH dates are logged on that row — used to show how long
    after first contact the follow-up happened).
"""
import io
import json
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

WORKBOOK_1_ID = "1EqHza6vadQoyrvNUnmK30s7u6LvcEI54Sy2YkFV5xoQ"  # PATHFINDER INQUIRIES
WORKBOOK_2_ID = "1ZSzNWIyoHA7UVRDA2cH9uVuhFa5uvhEmeBZifgokW74"  # Corporate/School AI Workshop Inquiries

# Sri Lanka Standard Time (UTC+5:30, no DST) — the "Last updated" stamp is shown in this
# timezone rather than UTC, since a UTC date can lag the team's actual local date by
# several hours (e.g. still showing "yesterday" until 5:30am local time).
LOCAL_TZ = timezone(timedelta(hours=5, minutes=30))

YES = "Yes"
NO = "No"
WHATSAPP = "Couldn't Connect, Sent Whatsapp"

# Column positions verified directly against the live workbooks (0-indexed).
# Header labels in a couple of these sheets are shifted relative to the real
# data by one column — these positions were confirmed against actual cell
# values, not the header text.
SHEET_CONFIGS = [
    dict(workbook=1, sheet="Analytics - With Colour Codes", display="Analytics",
         name_col=5, phone_col=6, contacted_col=8, followed_col=10,
         date_called_col=13, date_followed_col=14, track_dates=True,
         unique_col=20),
    dict(workbook=1, sheet="AI Program Inquiries (Apr 2026)", display="AI Program Inquiries (Apr 2026)",
         name_col=5, phone_col=6, contacted_col=8, followed_col=10,
         unique_col=18),
    dict(workbook=2, sheet="Corporates - June 2026", display="Corporates - June 2026",
         name_col=1, phone_col=2, contacted_col=5, followed_col=7,
         date_called_col=13, date_followed_col=14, track_dates=True,
         unique_col=15),
    dict(workbook=2, sheet="CorporateSchools - March 2026", display="Corporate/Schools - March 2026",
         name_col=4, phone_col=5, contacted_col=8, followed_col=10,
         date_called_col=14, date_followed_col=15, track_dates=True,
         unique_col=17),
    dict(workbook=1, sheet="Free AI Workshop for SchoolsIns", display="Free AI Workshop for Schools/Institutes",
         name_col=2, phone_col=3, contacted_col=6, followed_col=8,
         date_called_col=11, date_followed_col=12, track_dates=True,
         unique_col=16),
    # NOTE: the live tab name "Free AI Workshop For Jaffna Educators" is 37 characters,
    # over Excel's 31-character worksheet-name limit — Google Sheets silently truncates
    # it to 31 chars on .xlsx export (same reason "Free AI Workshop for SchoolsIns" above
    # is a truncated name, not the full tab name), so we must match on the truncated form.
    dict(workbook=1, sheet="Free AI Workshop For Jaffna Edu", display="Free AI Workshop For Jaffna Educators",
         name_col=1, phone_col=2, contacted_col=4, followed_col=6),
    # "Date Called" (col L) and "Date Followed Up" (col M) were added to this sheet
    # on 2026-08-07 specifically so this sheet could join the "Recent activity by
    # date" section, same as the other tracked sheets — the sheet had no date
    # columns at all before this.
    dict(workbook=2, sheet="Corporate - Aug 2026", display="Corporate - Aug 2026",
         name_col=3, phone_col=4, contacted_col=7, followed_col=9,
         date_called_col=11, date_followed_col=12, track_dates=True),
]


def download_workbook(spreadsheet_id):
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    return io.BytesIO(resp.content)


def fmt_date(value):
    """Normalize a raw date cell to 'D Mon' for display. Returns None if unparseable."""
    if pd.isna(value):
        return None
    try:
        dt = pd.to_datetime(value)
        return dt.strftime("%-d %b")
    except Exception:
        return None


def analyze_sheet(xls, cfg):
    df = pd.read_excel(xls, sheet_name=cfg["sheet"], header=None, dtype=str)
    data = df.iloc[1:]
    data = data[data[cfg["name_col"]].notna()]

    contacted_vals = data[cfg["contacted_col"]]
    conv = int((contacted_vals == YES).sum())
    whatsapp = int((contacted_vals == WHATSAPP).sum())
    contacted = conv + whatsapp
    total = int(len(data))
    not_contacted = total - contacted

    # Follow-up is a three-way classification (see module docstring): exactly "Yes" =
    # successful, exactly "No" = attempted-but-unsuccessful, blank = not attempted at all.
    # not_attempted is computed as a remainder so the three always reconcile to `total`.
    followed_vals = data[cfg["followed_col"]]
    follow_successful = int((followed_vals == YES).sum())
    follow_unsuccessful = int((followed_vals == NO).sum())
    follow_attempted = follow_successful + follow_unsuccessful
    # "Not followed up" is scoped to CONTACTED leads only — you can't follow up on someone
    # who was never reached in the first place, so leads that were never contacted are not
    # counted here at all (they simply aren't part of the follow-up conversation yet).
    # Clamped at 0 as a safety net: "contacted" and "followed up" are read from two
    # independently-maintained columns, so a data-entry inconsistency could in principle
    # show a follow-up attempt logged against a row that isn't exact-match "contacted".
    follow_not_attempted = max(0, contacted - follow_attempted)

    def _rate(n, d):
        return round(n / d, 4) if d else 0.0

    result = dict(
        name=cfg["display"], total=total, conv=conv, whatsapp=whatsapp,
        not_contacted=not_contacted, contacted=contacted,
        followed=follow_successful,  # backward-compat alias = follow_successful
        follow_attempted=follow_attempted,
        follow_successful=follow_successful,
        follow_unsuccessful=follow_unsuccessful,
        follow_not_attempted=follow_not_attempted,
        follow_attempt_rate=_rate(follow_attempted, total),
        follow_success_rate=_rate(follow_successful, total),
        follow_success_rate_of_attempts=_rate(follow_successful, follow_attempted),
        note=None,
    )
    if total and not follow_attempted:
        result["note"] = (
            "No follow-up attempts have been logged on this sheet yet — check whether these are "
            "recent leads or if follow-ups are overdue."
        )
    elif total and follow_attempted and not follow_successful:
        result["note"] = (
            "Follow-up attempts have been made on this sheet, but none have resulted in a "
            "successful conversation yet."
        )

    detail = None
    if cfg.get("track_dates"):
        rows = []
        for _, r in data.iterrows():
            called_raw = r[cfg["date_called_col"]]
            followed_raw = r[cfg["date_followed_col"]]
            called = fmt_date(called_raw)
            followed_d = fmt_date(followed_raw)
            if called or followed_d:
                name_val = r[cfg["name_col"]]
                phone_val = r[cfg["phone_col"]]

                # How this lead was reached, for the "contacted" list — reuses the same
                # exact-match contact-status column as the headline methodology above.
                contact_val = r[cfg["contacted_col"]]
                if contact_val == YES:
                    method = "Phone call"
                elif contact_val == WHATSAPP:
                    method = "WhatsApp"
                else:
                    method = ""

                # Days between first contact and follow-up, for the "followed up" list —
                # only computable when both dates are actually logged on this row.
                days_to_follow = None
                if called and followed_d:
                    try:
                        d1 = pd.to_datetime(called_raw).normalize()
                        d2 = pd.to_datetime(followed_raw).normalize()
                        days_to_follow = int((d2 - d1).days)
                    except Exception:
                        days_to_follow = None

                # Outcome of the follow-up attempt, for the "followed up" list — same
                # three-way exact-match read as the sheet-level follow-up breakdown above.
                # This is independent of whether a follow-up DATE was logged; a row can have
                # a logged follow-up date and still have been unsuccessful ("No").
                follow_status_val = r[cfg["followed_col"]]
                if follow_status_val == YES:
                    follow_status = "Successful"
                elif follow_status_val == NO:
                    follow_status = "Unsuccessful"
                else:
                    follow_status = ""

                rows.append(dict(
                    name=str(name_val).strip() if pd.notna(name_val) else "",
                    phone=str(phone_val).strip() if pd.notna(phone_val) else "",
                    called=called or "—",
                    followed=followed_d or "—",
                    sheet=cfg["display"],
                    workbook=cfg["workbook"],
                    method=method,
                    days_to_follow=days_to_follow,
                    follow_status=follow_status,
                ))
        detail = rows

    # Duplicate-row scan: every sheet has a final "Unique"-marker column, filled in
    # manually by the team for every row. It reads exactly "Unique" for normal rows;
    # anything else non-blank means that row has been flagged as a duplicate (the
    # flag text itself varies — we surface it as-is rather than assuming one format).
    dup_rows = []
    if cfg.get("unique_col") is not None:
        for _, r in data.iterrows():
            marker = r[cfg["unique_col"]]
            marker_str = str(marker).strip() if pd.notna(marker) else ""
            if marker_str and marker_str.lower() != "unique":
                name_val = r[cfg["name_col"]]
                phone_val = r[cfg["phone_col"]]
                dup_rows.append(dict(
                    name=str(name_val).strip() if pd.notna(name_val) else "",
                    phone=str(phone_val).strip() if pd.notna(phone_val) else "",
                    sheet=cfg["display"],
                    workbook=cfg["workbook"],
                    flag=marker_str,
                ))

    return result, detail, dup_rows


def bucket_by_date(detail, field, with_method=False):
    """Tally detail rows per date. When with_method=True (used for the "contacted, by
    date" list), also split each day's count into phone-call vs. WhatsApp using the
    same `method` field already computed per-row in analyze_sheet — so the dashboard
    can show a phone/WhatsApp percentage split per day, not just a raw daily count.
    `unknown` catches rows with a logged date but no exact-match method (e.g. contact
    status wasn't exactly "Yes"/WhatsApp-string even though a call date exists), so
    phone + whatsapp + unknown always reconciles back to the day's total `n`.
    """
    counts = {}
    method_counts = {}
    for r in detail:
        d = r[field]
        if d == "—":
            continue
        counts[d] = counts.get(d, 0) + 1
        if with_method:
            bucket = method_counts.setdefault(d, {"phone": 0, "whatsapp": 0})
            if r.get("method") == "Phone call":
                bucket["phone"] += 1
            elif r.get("method") == "WhatsApp":
                bucket["whatsapp"] += 1

    def sort_key(item):
        # NOTE: `item` here is a dict ({"date": ..., "n": ...}), not a tuple — must
        # index with item["date"], NOT item[0] (a previous version used item[0],
        # which raised KeyError on every call since dicts don't support positional
        # indexing; the exception was silently caught and fell back to datetime.max
        # for every entry, which made sorted() a no-op that just preserved whatever
        # order the raw sheet rows happened to be scanned in — not chronological
        # order at all. Fixed 2026-08-06.)
        try:
            return datetime.strptime(f"{item['date']} {datetime.now().year}", "%d %b %Y")
        except Exception:
            return datetime.max

    result = []
    for d, n in counts.items():
        entry = dict(date=d, n=n)
        if with_method:
            mc = method_counts.get(d, {"phone": 0, "whatsapp": 0})
            entry["phone"] = mc["phone"]
            entry["whatsapp"] = mc["whatsapp"]
            entry["unknown"] = n - mc["phone"] - mc["whatsapp"]
        result.append(entry)

    return sorted(result, key=sort_key)


def main():
    wb1 = download_workbook(WORKBOOK_1_ID)
    wb2 = download_workbook(WORKBOOK_2_ID)
    xls1 = pd.ExcelFile(wb1)
    xls2 = pd.ExcelFile(wb2)

    sheets = []
    all_detail = []
    all_duplicates = []
    for cfg in SHEET_CONFIGS:
        xls = xls1 if cfg["workbook"] == 1 else xls2
        result, detail, dup_rows = analyze_sheet(xls, cfg)
        sheets.append(result)
        if detail is not None:
            all_detail.extend(detail)
        all_duplicates.extend(dup_rows)

    recent_activity = None
    if all_detail:
        tracked_sheets = [cfg["display"] for cfg in SHEET_CONFIGS if cfg.get("track_dates")]
        tracked_contacted = sum(s["contacted"] for s in sheets if s["name"] in tracked_sheets)
        tracked_followed = sum(s["followed"] for s in sheets if s["name"] in tracked_sheets)
        recent_activity = dict(
            sheets=tracked_sheets,
            total_sheets=len(SHEET_CONFIGS),
            contacted_total=tracked_contacted,
            followed_total=tracked_followed,
            called_by_date=bucket_by_date(all_detail, "called", with_method=True),
            followed_by_date=bucket_by_date(all_detail, "followed"),
            detail=all_detail,
        )

    data = dict(
        updated=datetime.now(LOCAL_TZ).strftime("%-d %b %Y"),
        sheets=sheets,
        recent_activity=recent_activity,
        duplicates=all_duplicates,
    )

    with open("template.html", "r", encoding="utf-8") as f:
        template = f.read()

    data_json = json.dumps(data, indent=2, ensure_ascii=False)
    output = template.replace("/*__DATA_JSON__*/", data_json)

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(output)

    print(f"Wrote index.html — {sum(s['total'] for s in sheets)} total leads across {len(sheets)} sheets.")


if __name__ == "__main__":
    main()
