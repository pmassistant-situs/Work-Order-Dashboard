#!/usr/bin/env python3
"""
Work Order Dashboard - Automated Refresh Script (v5 — GitHub Actions)
Pulls the daily work-order workbook directly from Google Drive (shared as
"anyone with the link can view") via Google's public XLSX export endpoint —
no OAuth/service account needed.

Reads directly from the 'AppFolioReport' tab (the raw AppFolio Report
Builder export) inside the 'Work Orders - Daily Summary' workbook, rather
than the derived 'PowerQueryresult' tab. This has two advantages:
  1. The AppFolio work order link is constructed directly from the
     Work Order ID + Service Request ID columns already present here —
     no dependency on a Power Query formula elsewhere in the workbook.
  2. If Vendor / Primary Resident Phone Number / Primary Resident Email
     are ever added as extra columns to that same Report Builder export,
     this script picks them up automatically — no second file needed.
"""

import os, math, tempfile, urllib.request

DRIVE_FILE_ID = os.environ.get("DRIVE_FILE_ID", "1xhjrKwC9bHKAdYDVfNj3TUfM09qT4RKv")
EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"

OPEN_STATUSES = {'Assigned', 'Scheduled'}
APPFOLIO_SUBDOMAIN = "situsgroup"  # https://situsgroup.appfolio.com/...

def download_sheet():
    print(f"  Downloading sheet export (file ID: {DRIVE_FILE_ID})")
    req = urllib.request.Request(EXPORT_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as r:
        data = r.read()
    if not data or len(data) < 1000 or not data.startswith(b"PK"):
        raise Exception(
            "Downloaded file doesn't look like a valid XLSX export. "
            "Check that the sheet is still shared as 'Anyone with the link can view' "
            f"and that the file ID ({DRIVE_FILE_ID}) is correct."
        )
    print(f"  Downloaded {len(data):,} bytes")
    return data

def clean(v):
    if v is None: return None
    if isinstance(v, float) and math.isnan(v): return None
    if hasattr(v, "isoformat"): return str(v)[:10]
    s = str(v).strip() if not isinstance(v, (int, float)) else v
    if isinstance(s, str) and s.lower() in ('nan', ''):
        return None
    return s

def clean_text(v, maxlen):
    if v is None:
        return None
    s = str(v).replace('_x000D_', ' ').replace('\r', ' ').replace('\n', ' ')
    s = ' '.join(s.split())
    return s[:maxlen].strip() or None

def build_url(work_order_id, service_request_id):
    if not work_order_id or not service_request_id:
        return None
    try:
        woid = int(float(work_order_id))
        srid = int(float(service_request_id))
    except (ValueError, TypeError):
        return None
    return f"https://{APPFOLIO_SUBDOMAIN}.appfolio.com/maintenance/service_requests/{srid}/work_orders/{woid}"

def process_xlsx(xlsx_bytes):
    import pandas as pd

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(xlsx_bytes)
        tmp_path = f.name

    xl = pd.ExcelFile(tmp_path)
    sheet = "AppFolioReport" if "AppFolioReport" in xl.sheet_names else xl.sheet_names[0]
    raw = pd.read_excel(tmp_path, sheet_name=sheet, header=None)
    print(f"  Reading sheet '{sheet}' — raw shape: {raw.shape}")

    # Detect header row dynamically (metadata rows at top vary in count)
    header_row = None
    for i in range(min(20, len(raw))):
        row_vals = [str(v).strip() for v in raw.iloc[i].tolist()]
        if 'Property' in row_vals and 'Work Order Number' in row_vals:
            header_row = i
            print(f"  Header detected at row {i}")
            break
    if header_row is None:
        raise Exception("Could not find header row (expected 'Property' + 'Work Order Number' columns)")

    df = pd.read_excel(tmp_path, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    has_vendor = "Vendor" in df.columns
    has_phone  = "Primary Resident Phone Number" in df.columns
    has_email  = "Primary Resident Email" in df.columns
    if not (has_vendor and has_phone and has_email):
        print("  ℹ️  Vendor/Phone/Email columns not present in this export — "
              "add them to the AppFolio Report Builder query to include them here.")

    records = []
    for _, row in df.iterrows():
        wo = clean(row.get("Work Order Number"))
        status = clean(row.get("Status"))
        if not wo or not status:
            continue
        if status not in OPEN_STATUSES:
            continue
        wo_type = clean(row.get("Work Order Type"))
        if wo_type == "Unit Turn":
            continue

        prop_abbrev = clean(row.get("PropertyAbbrev."))
        prop_full = clean(row.get("Property"))
        prop = prop_abbrev or (prop_full.split(' - ')[0].strip() if prop_full and ' - ' in prop_full else prop_full)

        r = {
            "Property":     prop,
            "Unit":         clean(row.get("Unit")),
            "Status":       status,
            "Priority":     clean(row.get("Priority")),
            "Type":         wo_type,
            "WONumber":     wo,
            "AssignedUser": clean(row.get("Assigned User")),
            "Vendor":       clean(row.get("Vendor")) if has_vendor else None,
            "CreatedAt":    str(row.get("Created At", ""))[:10] if row.get("Created At") else None,
            "Description":  clean_text(row.get("Service Request Description") or row.get("Job Description"), 400),
            "StatusNotes":  clean_text(row.get("Status Notes"), 300),
            "Phone":        clean(row.get("Primary Resident Phone Number")) if has_phone else None,
            "Email":        clean(row.get("Primary Resident Email")) if has_email else None,
            "URL":          build_url(row.get("Work Order ID"), row.get("Service Request ID")),
        }
        records.append(r)

    os.unlink(tmp_path)

    seen = {}
    for r in records:
        wo = r['WONumber']
        if wo not in seen:
            seen[wo] = r
        else:
            existing = seen[wo]['AssignedUser'] or ''
            new_user = r['AssignedUser'] or ''
            if new_user and new_user not in existing:
                seen[wo]['AssignedUser'] = (existing + ', ' + new_user).strip(', ')
    records = list(seen.values())

    with_url = sum(1 for r in records if r['URL'])
    print(f"  Valid records after dedup: {len(records)}  (with working link: {with_url})")
    return records

def build_dashboard(records, date_str):
    import json
    with open("dashboard_template.html") as f:
        template = f.read()
    html = template.replace("DATA_PLACEHOLDER", json.dumps(records))
    html = html.replace("DATE_PLACEHOLDER", date_str)
    html = html.replace("COUNT_PLACEHOLDER", str(len(records)))
    return html

if __name__ == "__main__":
    from datetime import datetime, timezone

    print("📥 Downloading sheet from Google Drive...")
    xlsx_bytes = download_sheet()

    print("⚙️  Processing data...")
    records = process_xlsx(xlsx_bytes)
    print(f"✓ {len(records)} work orders processed")

    date_str = datetime.now(timezone.utc).strftime("%B %-d, %Y")

    print("🏗️  Building dashboard...")
    html = build_dashboard(records, date_str)

    with open("index.html", "w") as f:
        f.write(html)
    print(f"✓ Dashboard written ({len(html):,} chars)")
    print("✅ Done!")
