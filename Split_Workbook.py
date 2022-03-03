from pathlib import Path  # Core Python Module

import pandas as pd
import xlwings as xw  # pip install xlwings

this_dir = Path(__file__).parent
output_dir = this_dir / "output"
excel_file = this_dir / "Financial_Sample.xlsx"

# create output dir if it does not exist
output_dir.mkdir(parents=True, exist_ok=True)

# Create unquie list with items for a specific column
df = pd.read_excel(excel_file)
unique_values = df["Country"].unique()

# open Excel in the background
with xw.App(visible=True) as app:
    # open workbook
    wb = app.books.open(excel_file)
    # manipulate first worksheet
    sht = wb.sheets[0]

    for unique_value in unique_values:
        # Apply Autofilter to current region, copy data and turn off AutoFilter
        sht.api.Range("A1").CurrentRegion.AutoFilter(
            Field := 2, Criteria := unique_value
        )
        sht.range("A1").current_region.copy()
        sht.api.AutoFilterMode = False

        # Create new workbook and paste data
        new_wb = xw.books.add()
        new_wb.sheets[0].range("A1").paste()

        # Save new workbook in the output dir and close it
        new_wb.save(output_dir / f"{unique_value}.xlsx")
        new_wb.close()
