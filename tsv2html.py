#!/usr/bin/env python3
"""
tsv_to_html.py — Convert a TSV (Tab-Separated Values) file into an HTML table.

Usage:
  python tsv_to_html.py input.tsv output.html
"""

import html
import sys


def tsv_to_html(tsv_file, html_file):
    with open(tsv_file, "r", encoding="utf-8") as f:
        lines = [line.strip().split("\t") for line in f if line.strip()]

    with open(html_file, "w", encoding="utf-8") as out:
        out.write("<!DOCTYPE html>\n<html>\n<head>\n")
        out.write("<meta charset='utf-8'>\n<title>TSV to HTML Table</title>\n")
        out.write("<style>\n")
        out.write("table { border-collapse: collapse; width: 100%; }\n")
        out.write(
            "th, td { border: 1px solid #ccc; padding: 6px; text-align: left; }\n"
        )
        out.write("th { background-color: #f4f4f4; }\n")
        out.write("</style>\n</head>\n<body>\n")
        out.write("<table>\n")

        # Write header row
        if lines:
            headers = lines[0]
            out.write("  <thead><tr>")
            for h in headers:
                out.write(f"<th>{html.escape(h)}</th>")
            out.write("</tr></thead>\n  <tbody>\n")

            # Write data rows
            for row in lines[1:]:
                out.write("    <tr>")
                for cell in row:
                    out.write(f"<td>{html.escape(cell)}</td>")
                out.write("</tr>\n")

            out.write("  </tbody>\n")

        out.write("</table>\n</body>\n</html>\n")

    print(f"✅ HTML table written to: {html_file}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python tsv_to_html.py input.tsv output.html")
        sys.exit(1)

    tsv_to_html(sys.argv[1], sys.argv[2])
