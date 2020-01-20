import json

from libs.config import open_workbook, get_client
from modules.headers_from_cell_parse import separate_headers_vals

workbook_name = "Test sheet"
start_at, num_rows = 2, 2
sheet_idx = 1

def fetch_headers_vals(workbook_name, sheet_idx, start_at, num_rows):

    wb = open_workbook(workbook_name)
    ws = wb.get_worksheet(sheet_idx)
    all_vals = ws.get_all_values()

    headers, vals = separate_headers_vals(all_vals, start_at, num_rows)

    return headers, vals

headers, vals = fetch_headers_vals(workbook_name, sheet_idx, start_at, num_rows)


####################################

client = get_client()

wbs = client.list_spreadsheet_files()
wbs_names = [i['name'] for i in wbs][1:]
wbs_open = [open_workbook(name) for name in wbs_names]
wbs_sheets = [ws.worksheets() for ws in wbs_open]

sheet_dict = {}
for wb in wbs_sheets:
    for i, sheet in enumerate(wb):
        entries_dict = {}
        entries_dict['sheet'] = sheet
        entries_dict['sheet_idx'] = i

        sheet_dict[sheet.title] = entries_dict


wbs_titles = [[ws.title for ws in wb] for wb in wbs_sheets]
wbs_sets = [set(i) for i in wbs_titles]

# make sheet_dict from wb_template.json somehow

wb = open_workbook(workbook_name)
def alt_fetch_headers_vals(wb, sheet_idx, start_at, num_rows):
    ws = wb.get_worksheet(sheet_idx)
    all_vals = ws.get_all_values()

    headers, vals = separate_headers_vals(all_vals, start_at, num_rows)

    return headers, vals

def add_headers_to_sheet_dict(sheet_dict):
    for i, sheet in enumerate(sheet_dict): 
        print(f"Fetching for {sheet}...") 
        start_at, num_rows = sheet_dict[sheet]['header_info_tuple'] 
        headers, _ = alt_fetch_headers_vals(wb, i, start_at, num_rows) 
        sheet_dict[sheet]['headers'] = headers 

    return sheet_dict

def build_sheet_dict(wb):
    sheet_dict = {}
    worksheets = wb.worksheets()
    for i, sheet in enumerate(worksheets): 
        entries_dict = {} 
        entries_dict['sheet'] = sheet 
        entries_dict['sheet_idx'] = i 

        sheet_dict[sheet.title] = entries_dict

    with open('wb_template.json', 'r') as f:
        contents = f.read()
        template_dict = json.loads(contents)

    for sheet in sheet_dict:
        sheet_dict[sheet]['header_info_tuple'] = template_dict[sheet]['header_info_tuple']

    return sheet_dict

def sheet_fetch_headers_vals(ws, start_at, num_rows):
    all_vals = ws.get_all_values()

    headers, vals = separate_headers_vals(all_vals, start_at, num_rows)

    return headers, vals

a = build_sheet_dict(wbs_open[0])
b = build_sheet_dict(wbs_open[1])
c = build_sheet_dict(wbs_open[2])
sheet_dicts = [a, b, c]

sheetnames = list(a.keys())

def fetch_header_compare(sheet_dicts, sheetname='Cover Sheets'):
    print(f"Fetching compare for {sheetname}...")
    header_compare = [] 
    for wb in sheet_dicts:
        ws = wb[sheetname]['sheet'] 
        start_at, num_rows = wb[sheetname]['header_info_tuple'] 
        headers, _ = sheet_fetch_headers_vals(ws, start_at, num_rows) 
        header_compare.append(headers)

    header_set = [set(i) for i in header_compare]

    return header_compare, header_set

header_compare, header_set = fetch_header_compare(sheet_dicts, sheetnames[0])
header_set[0] - header_set[1]
