from bs4 import BeautifulSoup

def sort_html_table_by_first_column(input_html, output_html):
    # Read the HTML file
    with open(input_html, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')

    # Find the first table in the HTML
    table = soup.find('table')
    if not table:
        print("No table found in the HTML.")
        return

    # Extract all rows
    rows = table.find_all('tr')

    # Separate header and data rows
    header = rows[0] if rows else None
    data_rows = rows[1:] if len(rows) > 1 else []

    # Sort data rows by the text in the first <td>
    sorted_rows = sorted(
        data_rows,
        key=lambda row: row.find_all('td')[0].get_text(strip=True).lower() if row.find_all('td') else ''
    )

    # Clear existing rows and rebuild the table
    table.clear()
    if header:
        table.append(header)
    for row in sorted_rows:
        table.append(row)

    # Write the sorted HTML back
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(str(soup.prettify()))

    print(f"Sorted table written to {output_html}")


# Example usage
if __name__ == "__main__":
    sort_html_table_by_first_column("searchcall.html", "sorted_output.html")
