import json
from typing import Any

from libs.types import SheetValues, Worksheet
from modules.headers_from_cell_parse import separate_headers_vals


def sheet_fetch_headers_vals(ws: Worksheet, start_at: int, num_rows: int) -> Any:
    all_vals: SheetValues = ws.get_all_values()

    headers, vals = separate_headers_vals(all_vals, start_at, num_rows)

    return headers, vals

def prepare_sheet_headers_fetch(wb_name, ws_name, sheet_dict):
    ws: Worksheet = sheet_dict[ws_name]['sheet']
    print(f"Fetching for '{ws_name}' | ", end="")
    print(f"In '{wb_name[:10]}...' workbook")
    h = sheet_dict[ws_name]['header_info_dict']
    start_at, num_rows = h['start_at'], h['num_rows']
    headers, _ = sheet_fetch_headers_vals(ws, start_at, num_rows)

    return ws_name, headers


def sheet_by_rows(headers, vals):
    sheet_by_rows = []
    for row in vals:
        # data = {' | '.join(headers[i]): val
        data = {json.dumps(headers[i]): val
                    for i, val in enumerate(row)}
        sheet_by_rows.append(data)

    return sheet_by_rows


def sheet_by_cols(headers, vals):
    sheet_by_cols = {json.dumps(header): [] for header in headers}
    for row in vals:
        for i, val in enumerate(row):
            keys = list(sheet_by_cols.keys())
            col_dict = sheet_by_cols[keys[i]]
            col_dict.append(val)

    return sheet_by_cols
