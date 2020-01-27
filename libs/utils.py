import json

from .config import open_workbook

def write_all_sheets_to_disk(workbook_name, client=None):
    sheets = open_workbook(workbook_name, client).worksheets()
    for i, sheet in enumerate(sheets):
        workbook = f"data/sheet{i+1}"
        with open(workbook, 'w') as f:
            all_vals = sheet.get_all_values()
            f.write(json.dumps(all_vals))

def make_headers(all_vals, starts_at=1, num_rows=1):
    start_idx = starts_at - 1
    assert 0 <= start_idx, "Min \"starts_at\" is 1"
    assert start_idx <= len(all_vals), \
        "\"start_idx\" longer than length of \"all_vals\""
    assert 0 <= start_idx + num_rows <= len(all_vals), \
        f"\"num_rows\" '{num_rows}' is either too big or less than zero ({start_idx}) | ({len(all_vals)}) | {all_vals}"

    end_idx = start_idx + num_rows
    headers = list(zip(*all_vals[start_idx:end_idx]))
    vals = all_vals[end_idx:]

    return headers, vals
