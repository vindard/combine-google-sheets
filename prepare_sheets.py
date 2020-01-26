import json
from typing import List, Dict, Sequence, Union, Any

from gspread.models import Spreadsheet, Worksheet

from libs.config import open_workbook, get_client
from modules.headers_from_cell_parse import separate_headers_vals

SHEET_DICT_TEMPLATE = 'wb_template.json'
MASTER_WORKBOOK = ''

# Custom types
# TODO fix SheetDict to stop throwing assignment mypy errors below
SheetDict = Dict[str, Union[int, str, dict, Sequence, Worksheet]]
SheetDicts = Dict[str, SheetDict]
Row = List[str]
SheetValues = List[Row]

def initialise_sheet_dicts(wbs_open: Dict[str, Spreadsheet]):
    wbs_sheets = {ws.title: ws.worksheets() for ws in wbs_open.values()}

    sheet_dicts = {}
    for name, wb in wbs_sheets.items():
        sheet_dict = {}
        for i, sheet in enumerate(wb):
            entries_dict = {}
            entries_dict['sheet'] = sheet
            entries_dict['sheet_idx'] = i

            sheet_dict[sheet.title] = entries_dict

        sheet_dicts[name] = sheet_dict

    return sheet_dicts

def fetch_all_workbooks() -> Dict[str, Spreadsheet]:
    client = get_client()

    print("Fetching all workbooks...")
    wbs = client.list_spreadsheet_files()
    wbs_names = [i['name'] for i in wbs][1:]
    wbs_open = {name: open_workbook(name) for name in wbs_names}
    print(f"Fetched {len(wbs_open)} workbooks from Google Sheets\n")

    return wbs_open

def build_sheet_dict(wb: Spreadsheet) -> SheetDict:
    print(f"Building sheet_dict for \'{wb.title}\'...")
    sheet_dict: SheetDict = {}
    worksheets: List[Worksheet] = wb.worksheets()
    for i, ws in enumerate(worksheets):
        entries_dict = {}
        entries_dict['sheet'] = ws
        entries_dict['sheet_idx'] = i

        sheet_dict[ws.title] = entries_dict

    with open(SHEET_DICT_TEMPLATE, 'r') as f:
        contents = f.read()
        template_dict: SheetDict = json.loads(contents)

    for sheet in sheet_dict:
        sheet_dict[sheet]['header_info_tuple'] = template_dict[sheet]['header_info_tuple']

    print(f"\'{wb.title}\' template built!\n")
    return sheet_dict

def wbs_to_sheet_dicts(wbs_open: Dict[str, Spreadsheet]) -> SheetDicts:
    print(f"Collecting sheet_dicts...\n")
    sheet_dicts = {wb_name: build_sheet_dict(wb)
                    for wb_name, wb in wbs_open.items()}

    print(f"{len(sheet_dicts)} sheet_dicts collected\n")
    return sheet_dicts


def sheet_fetch_headers_vals(ws: Worksheet, start_at: int, num_rows: int) -> Any:
    all_vals: SheetValues = ws.get_all_values()

    headers, vals = separate_headers_vals(all_vals, start_at, num_rows)

    return headers, vals

def add_headers_to_sheet_dict(sheet_dict: SheetDict) -> SheetDict:
    for i, ws_name in enumerate(sheet_dict):
        ws: Worksheet = sheet_dict[ws_name]['sheet']
        print(f"Fetching for {ws_name}...")
        start_at, num_rows = sheet_dict[ws_name]['header_info_tuple']
        headers, _ = sheet_fetch_headers_vals(ws, start_at, num_rows)
        sheet_dict[ws_name]['headers'] = headers

    return sheet_dict

def add_headers_to_sheet_dicts(sheet_dicts: SheetDicts) -> SheetDicts:
    updated_sheet_dicts = {name: add_headers_to_sheet_dict(sheet_dict)
                for name, sheet_dict in sheet_dicts.items()}

    return updated_sheet_dicts

def compare_sheet_headers(sheet_dicts, workbooks=(0, 1), sheet_num=4):
    wb1_idx, wb2_idx = workbooks
    sheet_idx = sheet_num - 1
    keys = list(sheet_dicts.keys())

    # Fetch ws1 dict for sheet_num
    wb1_name = keys[wb1_idx]
    wb1 = sheet_dicts[wb1_name]
    wb1_names = list(wb1.keys())

    wb1_compare_sheet_name = wb1_names[sheet_idx]
    wb1_compare_sheet = wb1[wb1_compare_sheet_name]
    wb1_sheet_headers = wb1_compare_sheet['headers']

    # Fetch ws2 dict for sheet_num
    wb2_name = keys[wb2_idx]
    wb2 = sheet_dicts[wb2_name]
    wb2_names = list(wb2.keys())

    wb2_compare_sheet_name = wb2_names[sheet_idx]
    wb2_compare_sheet = wb2[wb2_compare_sheet_name]
    wb2_sheet_headers = wb2_compare_sheet['headers']

    # Prepare results
    wb1_headers_tuple = wb1_sheet_headers, set(wb1_sheet_headers)
    wb2_headers_tuple = wb2_sheet_headers, set(wb2_sheet_headers)

    return wb1_headers_tuple, wb2_headers_tuple

def run():
    wbs_open = fetch_all_workbooks()
    sheet_dicts = wbs_to_sheet_dicts(wbs_open)
    sheet_dicts_w_headers = add_headers_to_sheet_dicts(sheet_dicts)

    return sheet_dicts_w_headers

def run_compare(sheet_dicts, sheet1_idx=0, sheet2_idx=1):
    print("\n====================")
    print(f"Running compare for workbooks '{sheet1_idx}' & '{sheet2_idx}'")
    for i in range(10):
        (h1, hs1), (h2, hs2) = compare_sheet_headers(sheet_dicts, (sheet1_idx, sheet2_idx), sheet_num=i)
        set_diff = hs1 - hs2
        print(f"\n\n[ {i} ]")
        if not set_diff:
            print("No differences!")
        else:
            for i, val in enumerate(set_diff):
                print(f"{i+1}. {val}")

