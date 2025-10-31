import pandas as pd


def html_table_to_spreadsheet(html_file, output_file="output.xlsx"):
    """
    Converts the first HTML table found in an HTML file into a spreadsheet (.xlsx).

    Args:
        html_file (str): Path to the HTML file containing the table.
        output_file (str): Path where the Excel file should be saved. Defaults to 'output.xlsx'.
    """
    try:
        # Read all tables from the HTML file
        tables = pd.read_html(html_file)

        if not tables:
            print("No tables found in the HTML file.")
            return

        # If multiple tables exist, save each in a separate Excel sheet
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            for i, table in enumerate(tables):
                sheet_name = f"Table_{i+1}"
                table.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Successfully converted {len(tables)} table(s) to '{output_file}'")

    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert HTML table(s) to spreadsheet")
    parser.add_argument("html_file", help="Path to the HTML file")
    parser.add_argument(
        "-o",
        "--output",
        default="output.xlsx",
        help="Output spreadsheet filename (default: output.xlsx)",
    )
    args = parser.parse_args()

    html_table_to_spreadsheet(args.html_file, args.output)
