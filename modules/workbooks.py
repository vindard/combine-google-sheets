from typing import List, Dict, Sequence, Tuple, Union, Any
from concurrent.futures import Future
from concurrent.futures.thread import ThreadPoolExecutor as PoolExecutor

from libs.config import open_workbook, get_client
from libs.types import Spreadsheet

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

    wbs_names = [i['name'] for i in wbs]
    wbs_open = threaded_open_workbooks(wbs_names)

    return wbs_open

