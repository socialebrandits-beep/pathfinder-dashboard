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
  - "Followed up" = the follow-up column reads EXACTLY "Yes".
  - not_followed  = contacted - followed
  - not_contacted = total - contacted
  - "total" = number of rows with a non-empty Name cell (trailing template
    rows in the sheet have no name and are excluded).
  - This rule is applied uniformly across all 5 sheets.
"""
import io
import json
from datetime import datetime, timezone

import pandas as pd
import requests

WORKBOOK_1_ID = "1EqHza6vadQoyrvNUnmK30s7u6LvcEI54Sy2YkFV5xoQ"  # PATHFINDER INQUIRIES
WORKBOOK_2_ID = "1ZSzNWIyoHA7UVRDA2cH9uVuhFa5uvhEmeBZifgokW74"  # Corporate/School AI Workshop Inquiries

YES = "Yes"
WHATSAPP = "Couldn't Connect, Sent Whatsapp"

# Column positions verified directly against the live workbooks (0-indexed).
# Header labels in a couple of these sheets are shifted relative to the real
# data by one column — these positions were confirmed against actual cell
# values, not the header text.
SHEET_CONFIGS = [
    dict(workbook=1, sheet="Analytics - With Colour Codes", display="Analytics",
         name_col=5, phone_col=6, contacted_col=8, followed_col=10,
         date_called_col=13, date_followed_col=14, track_dates=True),
    dict(workbook=1, sheet="AI Program Inquiries (Apr 2026)", display="AI Program Inquiries (Apr 2026)",
         name_col=5, phone_col=6, contacted_col=8, followed_col=10),
    dict(workbook=2, sheet="Corporates - June 2026", display="Corporates - June 2026",
         name_col=1, phone_col=2, contacted_col=5, followed_col=7,
         date_called_col=13, date_followed_col=14, track_dates=True),
    dict(workbook=2, sheet="CorporateSchools - March 2026", display="Corporate/Schools - March 2026",
         name_col=4, phone_col=5, contacted_col=8, followed_col=10,
         date_called_col=14, date_followed_col=15, track_dates=True),
    dict(workbook=1, sheet="Free AI Workshop for SchoolsIns", display="Free AI Workshop for Schools/Institutes",
         name_col=2, phone_col=3, contacted_col=6, followed_col=8,
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

    followed = int((data[cfg["followed_col"]] == YES).sum())

    result = dict(
        name=cfg["display"], total=total, conv=conv, whatsapp=whatsapp,
        not_contacted=not_contacted, contacted=contacted, followed=followed,
        note=None,
    )
    if total and not contacted:
        result["note"] = None
    if total and contacted and not followed:
        result["note"] = (
            "No leads on this sheet have been followed up yet — check whether these are "
            "recent leads or if follow-ups are overdue."
        )

    detail = None
    if cfg.get("track_dates"):
        rows = []
        for _, r in data.iterrows():
            called = fmt_date(r[cfg["date_called_col"]])
            followed_d = fmt_date(r[cfg["date_followed_col"]])
            if called or followed_d:
                name_val = r[cfg["name_col"]]
                phone_val = r[cfg["phone_col"]]
                rows.append(dict(
                    name=str(name_val).strip() if pd.notna(name_val) else "",
                    phone=str(phone_val).strip() if pd.notna(phone_val) else "",
                    called=called or "—",
                    followed=followed_d or "—",
                    sheet=cfg["display"],
                    workbook=cfg["workbook"],
                ))
        detail = rows
    return result, detail


def bucket_by_date(detail, field):
    counts = {}
    for r in detail:
        d = r[field]
        if d == "—":
            continue
        counts[d] = counts.get(d, 0) + 1

    def sort_key(item):
        try:
            return datetime.strptime(f"{item[0]} {datetime.now().year}", "%d %b %Y")
        except Exception:
            return datetime.max

    return [dict(date=d, n=n) for d, n in sorted(counts.items(), key=sort_key)]


def main():
    wb1 = download_workbook(WORKBOOK_1_ID)
    wb2 = download_workbook(WORKBOOK_2_ID)
    xls1 = pd.ExcelFile(wb1)
    xls2 = pd.ExcelFile(wb2)

    sheets = []
    all_detail = []
    for cfg in SHEET_CONFIGS:
        xls = xls1 if cfg["workbook"] == 1 else xls2
        result, detail = analyze_sheet(xls, cfg)
        sheets.append(result)
        if detail is not None:
            all_detail.extend(detail)

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
            called_by_date=bucket_by_date(all_detail, "called"),
            followed_by_date=bucket_by_date(all_detail, "followed"),
            detail=all_detail,
        )

    data = dict(
        updated=datetime.now(timezone.utc).strftime("%-d %b %Y"),
        sheets=sheets,
        recent_activity=recent_activity,
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
