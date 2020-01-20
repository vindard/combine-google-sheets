import json

from libs.config import open_workbook

def fetch_worksheet_data(workbook_name, client=None, sheet_num=1, from_web=False):
    if from_web:
        print("Fetching data from web...")
        sheet_index = sheet_num-1
        sheet = open_workbook(workbook_name, client).worksheets()[sheet_index]
        all_vals = sheet.get_all_values()

    else:
        print("Fetching data from disk...")
        all_vals = None
        try:
            with open(f'data/sheet{sheet_num}', 'r') as f:
                all_vals = json.loads(f.read())
        except Exception as e:
            print(e)

    return all_vals


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
