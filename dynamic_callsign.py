#!/usr/bin/env python3
"""
Read data from a Google Sheet (using a service account) and generate a
simple HTML table from it.

Setup:
    pip install google-api-python-client google-auth

Usage:
    python dynamic_callsign.py 


"""

import argparse
import html
import sys
import os
import json
import gspread
from datetime import datetime
from string import Template
from pathlib import Path

from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials


SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def fetch_values():
    jsoncontent = os.environ.get('CALLSIGN')
    spreadsheet_id = os.environ.get('PAGEID')
    # creds_json = "C:\\Users\\Jagadeesh\\Downloads\\eminent-yen-504603-j7-c2409eef7f40.json" 
    #with open(creds_json) as f:
    #    jsoncontent = f.read()
    creds_dict = json.loads(jsoncontent)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)

    
    sheet = client.open_by_key(spreadsheet_id).worksheet("hambel")

    result = sheet.get_all_values()

    return result


def rows_to_html(rows, has_header=True):
    if not rows:
        return "<p>No data found.</p>"

    # Normalize row lengths so the table doesn't have ragged columns
    max_cols = max(len(r) for r in rows)
    rows = [r + [""] * (max_cols - len(r)) for r in rows]

    lines = [] 

    body = rows

    for row in body:
        lines.append("    <tr>")
        for cell in row:
            lines.append(f"      <td>{html.escape(str(cell))}</td>")
        lines.append("    </tr>")
    return "\n".join(lines)


def build_html_page(table_html: str, title: str = "Spreadsheet Data") -> str:
    template_path = Path("header.html")
    raw_template_text = template_path.read_text(encoding="utf-8")
    html_template = Template(raw_template_text)
    data_to_inject = {
        "latest_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "html_table": table_html
    }
    html_output = html_template.safe_substitute(data_to_inject)
    return html_output


def main():
    try:
        rows = fetch_values()
    except Exception as e:
        print(f"Error fetching spreadsheet data: {e}", file=sys.stderr)
        sys.exit(1)

    table_html = rows_to_html(rows, has_header="")
    page_html = build_html_page(table_html, title="test")

    with open("callsign.html", "w", encoding="utf-8") as f:
        f.write(page_html)

    print(f"Wrote {len(rows)} row(s) to callsign.html")


if __name__ == "__main__":
    main()