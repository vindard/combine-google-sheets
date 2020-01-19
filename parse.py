import json
import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials

WORKBOOK_NAME = os.environ.get('WORKBOOK')
CREDS_JSON = os.environ['CREDS_JSON']


def get_client(creds_json=CREDS_JSON):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scope)
    client = gspread.authorize(creds)

    return client


def open_workbook(workbook_name, client=None):
    if not client:
        client = get_client()

    open_workbook = client.open(workbook_name)

    return open_workbook


def write_all_sheets_to_disk(workbook_name, client=None):
    sheets = open_workbook(workbook_name, client).worksheets()
    for i, sheet in enumerate(sheets): 
        workbook = f"data/sheet{i+1}" 
        with open(workbook, 'w') as f: 
            all_vals = sheet.get_all_values() 
            f.write(json.dumps(all_vals))


def fetch_worksheet_data(workbook_name, client=None, sheet_num=9, from_web=False):
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


def make_headers(all_vals, starts_at=2, num_rows=2):
    starts_at -= 1
    assert 0 <= starts_at, "Min \"starts_at\" is 1"
    assert starts_at <= len(all_vals), \
        "\"starts_at\" longer than length of \"all_vals\""
    assert 0 <= starts_at + num_rows <= len(all_vals), \
        "\"num_rows\" is either too big or less than zero"

    header_end = starts_at + num_rows
    headers = [i for i in zip(*all_vals[starts_at:header_end])]
    rows = all_vals[header_end:]

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

    return new_headers, rows


def sheet_by_rows(headers, rows):
    sheet_by_rows = []
    for row in rows:
        # data = {' | '.join(headers[i]): val
        data = {json.dumps(headers[i]): val
                    for i, val in enumerate(row)}
        sheet_by_rows.append(data)

    return sheet_by_rows


def sheet_by_cols(headers, rows):
    sheet_by_cols = {json.dumps(header): [] for header in headers}
    for row in rows: 
        for i, val in enumerate(row): 
            keys = list(sheet_by_cols.keys())
            col_dict = sheet_by_cols[keys[i]]
            col_dict.append(val)

    return sheet_by_cols



if __name__ == "__main__":
    all_vals = fetch_worksheet_data(
        workbook_name=WORKBOOK_NAME, 
        sheet_num=9,
        from_web=False
    )
    headers, rows = make_headers(
        all_vals,
        starts_at=2,
        num_rows=3
    )
    sheet_by_rows = sheet_by_rows(headers, rows)
    sheet_by_cols = sheet_by_cols(headers, rows)

    print(json.dumps(sheet_by_rows, indent=2))
    print(json.dumps(sheet_by_cols, indent=2))
