import openpyxl
from html import escape

# Path to your spreadsheet
input_file = "callsigns.xlsx"  # 👈 change this to your XLSX file path
input_file = "test.xlsx"  # 👈 change this to your XLSX file path

# Load workbook and sheet
wb = openpyxl.load_workbook(input_file)
sheet = wb.active

# Read headers
headers = [cell.value for cell in sheet[1]]

# Start HTML
html = "<table border='1' cellspacing='0' cellpadding='5'>\n"
html += "  <tr>\n"
for header in headers:
    html += f"    <th>{escape(str(header))}</th>\n"
html += "  </tr>\n"

# Read data rows
for row in sheet.iter_rows(min_row=2, values_only=True):
    html += "  <tr>\n"
    for cell in row:
        cell_text = "" if cell is None else escape(str(cell))
        html += f"    <td>{cell_text}</td>\n"
    html += "  </tr>\n"

html += "</table>"

# Save HTML to file
output_file = "callsigns_table.html"
with open(output_file, "w", encoding="utf-8") as f:
    f.write(html)

print(f"HTML table generated and saved to: {output_file}")
