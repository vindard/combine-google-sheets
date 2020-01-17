import json
import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials

WORKBOOK_NAME = os.environ['WORKBOOK']
CREDS_JSON = os.environ['CREDS_JSON']


def open_workbook(workbook_name, creds_json):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scope)
    client = gspread.authorize(creds)

    open_workbook = client.open(workbook_name)

    return open_workbook


def write_all_sheets_to_disk(workbook_name, creds_json):
    sheets = open_workbook(workbook_name, creds_json).worksheets()
    for i, sheet in enumerate(sheets): 
        workbook = f"data/sheet{i+1}" 
        with open(workbook, 'w') as f: 
            all_vals = sheet.get_all_values() 
            f.write(json.dumps(all_vals))


def fetch_sheet_data(workbook_name, creds_json, sheet_num=9, from_web=False):
    if from_web:
        print("Fetching data from web...")
        sheet_index = sheet_num-1
        sheet = open_workbook(workbook_name, creds_json).worksheets()[sheet_index]
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


def make_headers(starts_at=2, ends_at=4, all_vals=None):
    starts_at -= 1
    ends_at -= 1
    num_rows = ends_at - (starts_at + 1)

    headers = [i for i in zip(*all_vals[starts_at:ends_at+1])]
    vals = all_vals[ends_at+1:]

    new_headers = []
    curr = [''] * num_rows
    for col in headers:
        new_col = []
        for i, item in enumerate(col):
            if i == 0 and item:
                curr = list(col)
            new_col.append(item or curr[i])
            curr[i] = item or curr[i]
        new_headers.append(new_col)

    return new_headers, vals


def join_headers_vals(headers, vals):
    all_data = []
    for row in vals:
        data = {' | '.join(headers[i]): val
                    for i, val in enumerate(row)}
        all_data.append(data)

    return all_data



if __name__ == "__main__":
    all_vals = fetch_sheet_data(
        workbook_name=WORKBOOK_NAME, 
        creds_json=CREDS_JSON, 
        sheet_num=9,
        from_web=False
    )
    headers, vals = make_headers(all_vals=all_vals)
    all_data = join_headers_vals(headers, vals)

    print(json.dumps(all_data, indent=2))
