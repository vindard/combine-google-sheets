import json
from typing import List, Dict, Sequence, Tuple, Union, Any
from concurrent.futures import Future
from concurrent.futures.thread import ThreadPoolExecutor as PoolExecutor

from gspread.models import Spreadsheet, Worksheet

from libs.config import MASTER_WORKBOOK, SHEET_DICT_TEMPLATE
from libs.types import SheetDict, SheetDicts, Row, SheetValues
from modules.sheets import prepare_sheet_headers_fetch
from modules.workbooks import fetch_all_workbooks


def build_sheet_dict(wb_name: str, wb: Spreadsheet) -> Tuple[str, SheetDict]:
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


def run():
    wbs_open = fetch_all_workbooks()
    sheet_dicts = threaded_wbs_to_sheet_dicts(wbs_open)
    sheet_dicts_w_headers = threaded_add_headers_to_sheet_dicts(sheet_dicts)

    return sheet_dicts_w_headers


if __name__ == "__main__":
    sheet_dicts = run()
