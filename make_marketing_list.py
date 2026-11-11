"""
One-click marketing list builder
--------------------------------
Turns a raw HubSpot contacts export into a clean, ready-to-send marketing list
with a summary sheet (pie + bar chart) in one run.

Usage:
    python make_marketing_list.py sample_hubspot_contacts_export.csv
    python make_marketing_list.py export.csv --country Canada --industry "Financial Services"
    python make_marketing_list.py export.csv --title cfo,finance --stage customer,opportunity

The default criteria below were agreed with the marketing team once, so most
requests need no setup. Extra filters are optional and stack on top.
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---- 1. Agreed default criteria (edit once, reuse every time) -------------
DEFAULT_CRITERIA = {
    "lifecycle_stages": ["lead", "marketingqualifiedlead", "salesqualifiedlead", "opportunity", "customer"],
    "exclude_unsubscribed": True,
    "require_email": True,
}

STAGE_LABELS = {
    "subscriber": "Subscriber", "lead": "Lead", "marketingqualifiedlead": "MQL",
    "salesqualifiedlead": "SQL", "opportunity": "Opportunity", "customer": "Customer",
    "evangelist": "Evangelist", "other": "Other",
}
COUNTRY_FIX = {"canada": "Canada", "ca": "Canada", "united states": "United States", "usa": "United States",
               "us": "United States", "united kingdom": "United Kingdom", "uk": "United Kingdom"}
TITLE_FIX = {"cfo": "CFO", "ceo": "CEO", "coo": "COO", "vp": "VP", "hr": "HR", "it": "IT"}

OUTPUT_COLUMNS = ["First Name", "Last Name", "Email", "Job Title", "Company", "Industry",
                  "City", "Country", "Stage", "Owner", "Last Activity"]


def tidy_text(s):
    s = re.sub(r"\s+", " ", str(s)).strip()
    if not s or s.lower() == "nan":
        return ""
    words = [TITLE_FIX.get(w.lower(), w.capitalize() if (w.isupper() or w.islower()) else w) for w in s.split(" ")]
    words = [w.lower() if i and w.lower() in {"of", "and", "the", "for"} else w for i, w in enumerate(words)]
    return " ".join(words)


def clean(df):
    df = df.copy()
    for col in ["First Name", "Last Name", "Job Title", "Company Name", "City"]:
        df[col] = df[col].map(tidy_text)
    df["Email"] = df["Email"].fillna("").astype(str).str.strip().str.lower()
    df["Country/Region"] = df["Country/Region"].fillna("").astype(str).str.strip().str.lower().map(COUNTRY_FIX).fillna("")
    df["Industry"] = df["Industry"].fillna("Unknown").replace("", "Unknown")
    df["Lifecycle Stage"] = df["Lifecycle Stage"].fillna("").str.strip().str.lower()
    df["Last Activity Date"] = pd.to_datetime(df["Last Activity Date"], errors="coerce")
    # Keep the most recently active record when the same email appears twice
    df = df.sort_values("Last Activity Date", ascending=False, na_position="last")
    has_email = df["Email"] != ""
    df = pd.concat([df[has_email].drop_duplicates("Email"), df[~has_email]])
    return df


def apply_criteria(df, args):
    c = DEFAULT_CRITERIA
    mask = df["Lifecycle Stage"].isin(c["lifecycle_stages"])
    if c["require_email"]:
        mask &= df["Email"] != ""
    if c["exclude_unsubscribed"]:
        mask &= df["Marketing emails unsubscribed"].astype(str).str.lower() != "true"
    applied = [f"Stages: {', '.join(STAGE_LABELS[s] for s in c['lifecycle_stages'])}",
               "Has email", "Not unsubscribed"]
    # Optional extra criteria from marketing
    if args.country:
        vals = [v.strip().lower() for v in args.country.split(",")]
        mask &= df["Country/Region"].str.lower().isin(vals)
        applied.append(f"Country: {args.country}")
    if args.industry:
        vals = [v.strip().lower() for v in args.industry.split(",")]
        mask &= df["Industry"].str.lower().isin(vals)
        applied.append(f"Industry: {args.industry}")
    if args.title:
        pattern = "|".join(re.escape(v.strip()) for v in args.title.split(","))
        mask &= df["Job Title"].str.contains(pattern, case=False, na=False)
        applied.append(f"Job title contains: {args.title}")
    if args.stage:
        vals = [v.strip().lower() for v in args.stage.split(",")]
        mask &= df["Lifecycle Stage"].isin(vals)
        applied.append(f"Stage only: {args.stage}")
    return df[mask], applied


def to_output(df):
    out = pd.DataFrame({
        "First Name": df["First Name"], "Last Name": df["Last Name"], "Email": df["Email"],
        "Job Title": df["Job Title"], "Company": df["Company Name"], "Industry": df["Industry"],
        "City": df["City"], "Country": df["Country/Region"],
        "Stage": df["Lifecycle Stage"].map(STAGE_LABELS),
        "Owner": df["Contact owner"].fillna("Unassigned").replace("", "Unassigned"),
        "Last Activity": df["Last Activity Date"].dt.strftime("%d %b %Y").fillna("—"),
    })
    return out.sort_values(["Company", "Last Name"]).reset_index(drop=True)


HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def write_excel(out, raw_count, applied, path):
    wb = Workbook()

    # Sheet 1: Summary
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Marketing Contact List — Summary"
    ws["A1"].font = Font(bold=True, size=16, color="1F3A5F")
    ws["A2"] = f"Generated {datetime.now():%d %b %Y %H:%M}  ·  {len(out):,} contacts (from {raw_count:,} raw rows)"
    ws["A2"].font = Font(italic=True, color="666666")
    ws["A3"] = "Criteria: " + "  |  ".join(applied)
    ws["A3"].font = Font(color="666666")

    by_stage = out["Stage"].value_counts()
    by_ind = out["Industry"].value_counts()

    def table(start_row, col, title, series):
        ws.cell(start_row, col, title).font = HEADER_FONT
        ws.cell(start_row, col).fill = HEADER_FILL
        ws.cell(start_row, col + 1, "Contacts").font = HEADER_FONT
        ws.cell(start_row, col + 1).fill = HEADER_FILL
        for i, (k, v) in enumerate(series.items(), start=1):
            ws.cell(start_row + i, col, k)
            ws.cell(start_row + i, col + 1, int(v))
        return start_row + len(series)

    end_stage = table(5, 1, "Stage", by_stage)
    end_ind = table(5, 4, "Industry", by_ind)

    pie = PieChart()
    pie.title = "Contacts by Stage"
    pie.add_data(Reference(ws, min_col=2, min_row=5, max_row=end_stage), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=6, max_row=end_stage))
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    pie.dataLabels.showSerName = False
    pie.dataLabels.showLeaderLines = False
    pie.height, pie.width = 8, 12
    ws.add_chart(pie, "H2")

    bar = BarChart()
    bar.type = "bar"
    bar.title = "Contacts by Industry"
    bar.legend = None
    bar.add_data(Reference(ws, min_col=5, min_row=5, max_row=end_ind), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=4, min_row=6, max_row=end_ind))
    bar.x_axis.scaling.orientation = "maxMin"
    bar.height, bar.width = 9, 12
    ws.add_chart(bar, "H19")

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["D"].width = 22

    # Sheet 2: The list itself
    wl = wb.create_sheet("Contact List")
    wl.append(OUTPUT_COLUMNS)
    for cell in wl[1]:
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(vertical="center")
    for row in out.itertuples(index=False):
        wl.append(list(row))
    widths = [13, 14, 38, 24, 26, 20, 13, 15, 12, 12, 14]
    for i, w in enumerate(widths, start=1):
        wl.column_dimensions[get_column_letter(i)].width = w
    wl.freeze_panes = "A2"
    wl.auto_filter.ref = wl.dimensions

    wb.save(path)


def main():
    p = argparse.ArgumentParser(description="One-click marketing list from a HubSpot export")
    p.add_argument("export_csv")
    p.add_argument("--country", help="e.g. Canada or 'Canada,United States'")
    p.add_argument("--industry", help="e.g. 'Financial Services,Energy'")
    p.add_argument("--title", help="keywords, e.g. cfo,finance")
    p.add_argument("--stage", help="e.g. customer,opportunity")
    p.add_argument("--out", help="output .xlsx path")
    args = p.parse_args()

    raw = pd.read_csv(args.export_csv, dtype=str)
    df = clean(raw)
    filtered, applied = apply_criteria(df, args)
    if filtered.empty:
        sys.exit("No contacts match these criteria.")
    out = to_output(filtered)
    path = args.out or f"marketing_list_{datetime.now():%Y%m%d}.xlsx"
    write_excel(out, len(raw), applied, path)
    print(f"Done: {len(raw):,} raw rows -> {len(out):,} contacts -> {Path(path).name}")


if __name__ == "__main__":
    main()
