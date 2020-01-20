import json
import os

from modules.sheet import fetch_worksheet_data, sheet_by_cols, sheet_by_rows
# from modules.headers_from_cell_parse import separate_headers_vals
from modules.headers_from_merge_data import separate_headers_vals

WORKBOOK_NAME = os.environ.get('WORKBOOK')


if __name__ == "__main__":
    all_vals = fetch_worksheet_data(
        workbook_name=WORKBOOK_NAME, 
        sheet_num=9,
        from_web=True
    )
    headers, vals = separate_headers_vals(
        all_vals,
        starts_at=2,
        num_rows=3
    )
    breakpoint()
    sheet_by_rows = sheet_by_rows(headers, vals)
    sheet_by_cols = sheet_by_cols(headers, vals)

    print(json.dumps(sheet_by_rows, indent=2))
    print(json.dumps(sheet_by_cols, indent=2))
