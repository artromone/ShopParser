import os
import json
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def adjust_column_widths(sheet):
    for column in sheet.columns:
        max_length = max(len(str(cell.value)) for cell in column)
        column_letter = get_column_letter(column[0].column)
        adjusted_width = (max_length + 2) * 1.2
        sheet.column_dimensions[column_letter].width = adjusted_width


def format_datetime(datetime_str):
    datetime_obj = datetime.datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")
    formatted_datetime = datetime_obj.strftime("%d.%m.%y-%H:%M")
    return formatted_datetime


def replace_shop_name(shop_name):
    replace_file_path = os.path.join("sys", "shopnamereplace.json")

    # Load data from shopnamereplace.json for shop name replacements
    with open(replace_file_path, "r", encoding="utf-8") as replace_file:
        replace_data = json.load(replace_file)

    # Replace shop name if there is a matching entry in shopnamereplace.json
    for entry in replace_data:
        if entry.get("original_name") == shop_name:
            return entry.get("short_name")

    return shop_name


def create_export_file(data):
    # Path to the output directory
    output_dir = "output/"

    # Load excluded shop names
    excluded_file_path = os.path.join("sys", "excluded.json")
    excluded_shops = []
    if os.path.exists(excluded_file_path):
        with open(excluded_file_path, "r", encoding="utf-8") as excluded_file:
            excluded_shops = json.load(excluded_file)

    # Create a new Excel workbook for export
    export_workbook = Workbook()

    # Remove the default "Sheet" that is automatically created
    default_sheet = export_workbook["Sheet"]
    export_workbook.remove(default_sheet)

    # Check if there is any data available
    if not data:
        # Create a single sheet named "No data" if there is no data
        export_sheet = export_workbook.active
        export_sheet.title = "No data"
    else:
        # Group the data by month
        data_by_month = {}
        for item in data:
            receipt = item.get("ticket", {}).get("document", {}).get("receipt", {})
            datetime_str = receipt.get("dateTime")
            shop_name = replace_shop_name(receipt.get("user"))
            if datetime_str and shop_name not in excluded_shops:
                datetime_obj = datetime.datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")
                month_year = datetime_obj.strftime("%b %Y")
                if month_year not in data_by_month:
                    data_by_month[month_year] = []
                data_by_month[month_year].append(item)

        # Create separate sheets for each month
        for month_year, month_data in data_by_month.items():
            export_sheet = export_workbook.create_sheet(title=month_year)

            # Set the column headers for the export sheet
            export_headers = ["Time", "Shop", "Product", "Price", "Quantity", "Sum"]
            export_sheet.append(export_headers)

            # Apply bold font style to column headers
            export_bold_font = Font(bold=True)
            for cell in export_sheet[1]:
                cell.font = export_bold_font

            # Write data to the export sheet
            for item in month_data:
                receipt = item.get("ticket", {}).get("document", {}).get("receipt", {})
                shop_name = replace_shop_name(receipt.get("user"))
                for product in receipt.get("items", []):
                    row = [
                        format_datetime(receipt.get("dateTime")), shop_name, product.get("name"),
                        product.get("price") / 100, product.get("quantity"), product.get("sum") / 100
                    ]
                    export_sheet.append(row)

            # Adjust column widths
            adjust_column_widths(export_sheet)

    # Save the export workbook as export.xlsx
    export_file = os.path.join(output_dir, "export.xlsx")
    export_workbook.save(export_file)

    print(f"Export file {export_file} created successfully.")


# Path to the input directory
input_dir = "input/"

# Get a list of JSON files in the directory
json_files = [file for file in os.listdir(input_dir) if file.endswith(".json")]

# Load excluded shop names
excluded_file_path = os.path.join("sys", "excluded.json")
excluded_shops = []
if os.path.exists(excluded_file_path):
    with open(excluded_file_path, "r", encoding="utf-8") as excluded_file:
        excluded_shops = json.load(excluded_file)

data = []  # Store the data from all JSON files

for json_file in json_files:
    # Load data from the JSON file
    with open(os.path.join(input_dir, json_file), "r", encoding="utf-8") as file:
        data.extend(json.load(file))

# Create the export file with separate sheets for each month
create_export_file(data)
