import json
from concurrent.futures import Future
from concurrent.futures.thread import ThreadPoolExecutor as PoolExecutor
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

def threaded_open_workbooks(wbs_names: List[str]) -> Dict[str, Spreadsheet]:
    print(f"Opening {len(wbs_names)} workbooks...")
    wbs_open_tasks= []
    with PoolExecutor() as executor:
        for wb_name in wbs_names:
            f: Future = executor.submit(open_workbook, wb_name)
            wbs_open_tasks.append(f)

    wbs_open = {}
    for i, task in enumerate(wbs_open_tasks):
        wb_name, open_wb = task.result()
        print(f"{i+1}. Fetched '{wb_name}'")
        wbs_open[wb_name] = open_wb

    print(f"\nOpened {len(wbs_open)} workbooks from Google Sheets\n")

    return wbs_open

def fetch_all_workbooks() -> Dict[str, Spreadsheet]:
    client = get_client()

    print("\n======================")
    print("Fetch All Workbooks")
    print("======================\n")

    print("Fetching...")
    wbs = client.list_spreadsheet_files()
    print(f"Fetched {len(wbs)} workbooks from Google Sheets\n")

    wbs_names = [i['name'] for i in wbs][1:]
    wbs_open = threaded_open_workbooks(wbs_names)

    return wbs_open

def build_sheet_dict(wb_name: str, wb: Spreadsheet) -> SheetDict:
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
        sheet_dict[sheet]['header_info_dict'] = template_dict[sheet]['header_info_dict']

    print(f"\'{wb.title}\' template built!")
    return wb_name, sheet_dict

def threaded_wbs_to_sheet_dicts(wbs_open: Dict[str, Spreadsheet]) -> SheetDicts:
    print(f"\n======================")
    print(f"Collecting Sheet Dicts")
    print(f"======================\n")

    sheet_dicts_tasks = []
    with PoolExecutor() as executor:
        for wb_name, wb in wbs_open.items():
            f: Future = executor.submit(build_sheet_dict, wb_name, wb)
            sheet_dicts_tasks.append(f)

    # Collect results from complete tasks
    sheet_dicts = {sheet_dict.result()[0]: sheet_dict.result()[1]
                        for sheet_dict in sheet_dicts_tasks}
    print(f"\nFinished collecting sheet_dicts: {len(sheet_dicts)} sheet_dicts collected\n")

    return sheet_dicts


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


def threaded_add_headers_to_sheet_dict(wb_name: str, sheet_dict: SheetDict) -> SheetDict:
    header_tasks = []
    with PoolExecutor() as executor:
        for ws_name in sheet_dict:
            f: Future = \
                executor.submit(prepare_sheet_headers_fetch, wb_name, ws_name, sheet_dict)
            header_tasks.append(f)

    # Collect results from complete tasks
    for task in header_tasks:
        ws_name, headers = task.result()
        sheet_dict[ws_name]['headers'] = headers

    return wb_name, sheet_dict

def threaded_add_headers_to_sheet_dicts(sheet_dicts: SheetDicts) -> SheetDicts:
    print("\n=======================================")
    print(f"Fetching Sheet Headers for {len(sheet_dicts)} workbooks")
    print("=======================================")


    sheet_dicts_update_tasks = []
    with PoolExecutor() as executor:
        for wb_name, sheet_dict in sheet_dicts.items():
            f: Future = executor.submit(threaded_add_headers_to_sheet_dict, wb_name, sheet_dict)
            sheet_dicts_update_tasks.append(f)

    # Collect results from complete tasks
    updated_sheet_dicts = {}
    for task in sheet_dicts_update_tasks:
        wb_name, sheet_dict = task.result()
        updated_sheet_dicts[wb_name] = sheet_dict

    return updated_sheet_dicts

def compare_sheet_headers(sheet_dicts, base_wb_name, compare_wb_name, sheet_num=4):
    sheet_idx = sheet_num - 1
    keys = list(sheet_dicts.keys())

    # Fetch ws1 dict for sheet_num
    wb1_name = base_wb_name
    wb1 = sheet_dicts[wb1_name]
    wb1_names = list(wb1.keys())

    wb1_compare_sheet_name = wb1_names[sheet_idx]
    wb1_compare_sheet = wb1[wb1_compare_sheet_name]
    wb1_sheet_headers = wb1_compare_sheet['headers']

    # Fetch ws2 dict for sheet_num
    wb2_name = compare_wb_name
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
    sheet_dicts = threaded_wbs_to_sheet_dicts(wbs_open)
    sheet_dicts_w_headers = threaded_add_headers_to_sheet_dicts(sheet_dicts)

    return sheet_dicts_w_headers

def run_compare(sheet_dicts, wb_idx=0):
    keys = list(sheet_dicts.keys())
    base_wb_name = MASTER_WORKBOOK or keys[0]
    compare_wb_name = keys[wb_idx]

    print("\n====================")
    print(f"Running compare for workbooks:\n- '{base_wb_name}'\n- '{compare_wb_name}'")
    for i in range(10):
        (h1, hs1), (h2, hs2) = compare_sheet_headers(sheet_dicts, base_wb_name, compare_wb_name, sheet_num=i)
        set_diff = hs1 - hs2
        print(f"\n\n[ {i} ]")
        if not set_diff:
            print("No differences!")
        else:
            for i, val in enumerate(set_diff):
                print(f"{i+1}. {val}")

if __name__ == "__main__":
    sheet_dicts = run()
    run_compare(sheet_dicts, 2)
