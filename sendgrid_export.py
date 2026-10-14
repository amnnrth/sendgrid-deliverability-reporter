import os
import json
from pathlib import Path
from datetime import date, datetime, time, timedelta

import requests
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

STATS_API_URL = "https://api.sendgrid.com/v3/stats"
MSG_API_URL = "https://api.sendgrid.com/v3/messages"
EVENTS_FILE = Path("sendgrid_events.json")
OUTPUT_DIR = Path("reports")

NAVY = "17365D"
BLUE = "2878B5"
GREEN = "218739"
RED = "C62828"
ORANGE = "E69138"
LIGHT_BLUE = "D9EAF7"
LIGHT_RED = "FCE4D6"


def ask_date(label, default):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        try:
            result = datetime.strptime(value, "%Y-%m-%d").date()
            if result > date.today():
                print("Date cannot be in the future.")
                continue
            return result
        except ValueError:
            print("Use YYYY-MM-DD format.")


def select_date_range():
    today = date.today()
    
    print("\n--- Select Reporting Period ---")
    print("1. Last 7 days (Recommended for individual message logs)")
    print("2. Last 30 days")
    print("3. This month")
    print("4. Custom date range (YYYY-MM-DD)")
    
    while True:
        choice = input("\nChoose an option (1-4) [1]: ").strip() or "1"
        
        if choice == "1":
            end = today
            start = end - timedelta(days=7)
            print(f"Selected: {start} to {end}")
            return start, end
            
        elif choice == "2":
            end = today
            start = end - timedelta(days=30)
            print(f"Selected: {start} to {end}")
            return start, end
            
        elif choice == "3":
            end = today
            start = date(today.year, today.month, 1)
            print(f"Selected: {start} to {end}")
            return start, end
            
        elif choice == "4":
            print("\nEnter custom dates:")
            start = ask_date("Start date", (today - timedelta(days=7)).isoformat())
            end = ask_date("End date", today.isoformat())
            if start > end:
                print("Error: Start date must be on or before end date. Try again.")
                continue
            return start, end
            
        else:
            print("Invalid choice. Please enter 1, 2, 3, or 4.")


def fetch_stats(api_key, start, end):
    headers = {"Authorization": f"Bearer {api_key}"}
    params = {
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "aggregated_by": "day"
    }
    print(f"\nQuerying SendGrid Stats API ({start} to {end})...")
    response = requests.get(STATS_API_URL, headers=headers, params=params, timeout=60)
    if response.status_code != 200:
        print(f"Stats API warning [{response.status_code}]: {response.text}")
        return []
    return response.json()


def fetch_messages_api(api_key, start, end):
    headers = {"Authorization": f"Bearer {api_key}"}
    query = f'last_event_time BETWEEN TIMESTAMP "{start}T00:00:00Z" AND TIMESTAMP "{end}T23:59:59Z"'
    print(f"Querying SendGrid Messages API for individual message logs...")
    try:
        response = requests.get(
            MSG_API_URL,
            headers=headers,
            params={"query": query, "limit": 1000},
            timeout=60,
        )
        if response.status_code == 200:
            msgs = response.json().get("messages", [])
            print(f"Messages API found {len(msgs)} individual records.")
            return msgs
        else:
            print(f"Messages API response [{response.status_code}]: {response.text}")
    except Exception as e:
        print(f"Messages API notice: {e}")
    return []


def load_local_events(start, end):
    if not EVENTS_FILE.exists():
        return []
    try:
        raw = json.loads(EVENTS_FILE.read_text(encoding="utf-8"))
        if isinstance(raw, dict):
            raw = raw.get("events", [])
        events = []
        start_ts = int(datetime.combine(start, time.min).timestamp())
        end_ts = int(datetime.combine(end + timedelta(days=1), time.min).timestamp())
        for event in raw:
            try:
                ts = int(event.get("timestamp", 0))
            except (TypeError, ValueError):
                continue
            if start_ts <= ts < end_ts:
                events.append(event)
        return events
    except Exception:
        return []


def parse_data(stats_data, messages_data, events_data):
    # 1. Parse Stats for Daily Metrics
    rows = []
    for entry in stats_data:
        day = entry.get("date")
        for stat_group in entry.get("stats", []):
            metrics = stat_group.get("metrics", {})
            row = {"Date": day}
            row.update(metrics)
            rows.append(row)
            
    stats_df = pd.DataFrame(rows)
    if not stats_df.empty:
        numeric_cols = stats_df.columns.difference(["Date"])
        stats_df = stats_df.groupby("Date")[list(numeric_cols)].sum().reset_index()

    # 2. Parse Individual Email Details (Messages API + Local Events)
    detail_records = []
    
    for m in messages_data:
        detail_records.append({
            "Date": str(m.get("last_event_time", ""))[:19].replace("T", " "),
            "Sender": m.get("from_email") or m.get("from"),
            "Recipient": m.get("to_email") or m.get("email"),
            "Status": m.get("status", "Unknown"),
            "Remarks / Reason": m.get("reason") or m.get("response") or "No remarks recorded"
        })

    for e in events_data:
        ts = int(e.get("timestamp", 0))
        dt_str = datetime.fromtimestamp(ts, tz=__import__("datetime").timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
        detail_records.append({
            "Date": dt_str,
            "Sender": e.get("from") or e.get("marketing_campaign_name") or "Unknown",
            "Recipient": e.get("email") or e.get("to") or "Unknown",
            "Status": e.get("event") or e.get("type") or "Unknown",
            "Remarks / Reason": e.get("reason") or e.get("status") or e.get("response") or "Webhook Event Recorded"
        })

    email_df = pd.DataFrame(detail_records)
    if not email_df.empty:
        email_df = email_df.drop_duplicates(subset=["Date", "Recipient", "Status"])

    # 3. Summaries for Dashboards and Charts
    if not stats_df.empty:
        daily = stats_df[["Date", "requests"]].rename(columns={"requests": "Events"})
        total_delivered = stats_df["delivered"].sum()
        total_bounces = stats_df["bounces"].sum()
        total_blocks = stats_df["blocks"].sum()
        total_deferred = stats_df["deferred"].sum()
    else:
        daily = pd.DataFrame(columns=["Date", "Events"])
        total_delivered = total_bounces = total_blocks = total_deferred = 0

    status_summary = pd.DataFrame({
        "Status": ["Delivered", "Bounced", "Blocked", "Deferred"],
        "Count": [total_delivered, total_bounces, total_blocks, total_deferred]
    })

    failure_summary = pd.DataFrame({
        "Failure type": ["Bounces", "Blocks", "Deferred"],
        "Events": [total_bounces, total_blocks, total_deferred]
    })

    failed_df = email_df[email_df["Status"].str.lower().isin(["bounced", "blocked", "deferred", "dropped", "failed", "not_delivered"])] if not email_df.empty else pd.DataFrame()

    return stats_df, email_df, failed_df, status_summary, failure_summary, daily


def style_sheet(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 30

    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = max(len(str(c.value or "")) for c in col[:200]) + 2
        ws.column_dimensions[letter].width = min(max(width, 12), 45)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def add_table(writer, name, df):
    if df.empty:
        df = pd.DataFrame({"Info": ["No message-level records available. (Note: SendGrid retains individual message activity API logs for 3-7 days; older historical dates rely on Stats aggregates)."]})
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    style_sheet(ws)


def main():
    api_key = os.getenv("SENDGRID_API_KEY")
    if not api_key:
        raise SystemExit("Set SENDGRID_API_KEY environment variable before running.")

    print("\n=== SendGrid Master Deliverability Reporter ===")
    start, end = select_date_range()

    stats_data = fetch_stats(api_key, start, end)
    messages_data = fetch_messages_api(api_key, start, end)
    events_data = load_local_events(start, end)

    stats_df, email_df, failed_df, status_summary, failure_summary, daily = parse_data(
        stats_data, messages_data, events_data
    )

    OUTPUT_DIR.mkdir(exist_ok=True)
    filename = OUTPUT_DIR / f"sendgrid_master_report_{start}_{end}.xlsx"
    if filename.exists():
        try:
            filename.rename(filename)
        except PermissionError:
            timestamp_suffix = datetime.now().strftime("%H%M%S")
            filename = OUTPUT_DIR / f"sendgrid_master_report_{start}_{end}_{timestamp_suffix}.xlsx"
            print(f"\nNotice: File locked in Excel. Saving as: {filename.name}")

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        wb = writer.book
        ws = wb.create_sheet("Executive Dashboard", 0)
        ws.sheet_view.showGridLines = False
        ws.merge_cells("A1:F1")
        ws["A1"] = "SENDGRID MASTER DELIVERABILITY REPORT"
        ws["A1"].font = Font(name="Arial", size=18, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
        ws["A1"].alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 40

        ws["A3"] = "Reporting period"
        ws["B3"] = f"{start} to {end}"
        ws["A4"] = "Total Requests (Stats)"
        ws["B4"] = int(stats_df["requests"].sum()) if not stats_df.empty else 0
        ws["A5"] = "Total Delivered"
        ws["B5"] = int(stats_df["delivered"].sum()) if not stats_df.empty else 0

        delivered = int(stats_df["delivered"].sum()) if not stats_df.empty else 0
        requests_total = int(stats_df["requests"].sum()) if not stats_df.empty else 0

        ws["A7"] = "Global Performance"
        ws["A8"] = "Delivered Messages"
        ws["B8"] = delivered
        ws["A9"] = "Bounces / Blocks"
        ws["B9"] = int(stats_df["bounces"].sum() + stats_df["blocks"].sum()) if not stats_df.empty else 0
        ws["A10"] = "Overall Delivery Rate %"
        ws["B10"] = round(delivered / requests_total * 100, 2) if requests_total else 0

        ws["D3"] = "Data Retention & Source Note"
        ws["D4"] = "Macro metrics & charts pull from SendGrid Stats API (/v3/stats) for full historical coverage."
        ws["D5"] = "Individual sender/recipient logs pull from Messages API (retained for 3-7 days by SendGrid)."

        for cell in ("A3", "A4", "A5", "A7", "A10", "D3"):
            ws[cell].font = Font(bold=True, color=NAVY)

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 4
        ws.column_dimensions["D"].width = 42

        add_table(writer, "Email Details", email_df)
        add_table(writer, "Failed Messages", failed_df)
        add_table(writer, "Daily Stats", stats_df)
        add_table(writer, "Failure Breakdown", failure_summary)
        add_table(writer, "Daily Events", daily)
        add_table(writer, "Status Summary", status_summary)

        if not daily.empty:
            chart = LineChart()
            chart.title = "Requests by Day (UTC)"
            chart.y_axis.title = "Request count"
            chart.x_axis.title = "Date"
            chart.height = 7
            chart.width = 14
            data = Reference(writer.sheets["Daily Events"], min_col=2, min_row=1, max_row=len(daily) + 1)
            cats = Reference(writer.sheets["Daily Events"], min_col=1, min_row=2, max_row=len(daily) + 1)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            ws.add_chart(chart, "A13")

        if not status_summary.empty:
            chart = BarChart()
            chart.type = "bar"
            chart.title = "Metrics by Status"
            chart.x_axis.title = "Count"
            chart.height = 7
            chart.width = 12
            data = Reference(writer.sheets["Status Summary"], min_col=2, min_row=1, max_row=len(status_summary) + 1)
            cats = Reference(writer.sheets["Status Summary"], min_col=1, min_row=2, max_row=len(status_summary) + 1)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            ws.add_chart(chart, "D13")

    print(f"\nExcel master report successfully created: {filename.resolve()}")


if __name__ == "__main__":
    main()