import pandas as pd
import json

def have_list(): 
    file_list = [
        '00-24-sheets.json',
        '25-48-sheets.json',
        '49-72-sheets.json',
        '72-96-sheets.json',
        '97-120-sheets.json',
        '121-144-sheets.json',
        '145-150-sheets.json',
    ]
    all = {}
    for filename in file_list:
        with open(filename, 'r') as f: 
            a = f.read() 
            a = json.loads(a)
            all = {**all, **a}
    return list(all.keys()) 

def print_differences(dfs_by_sheet):
    for ws_name, ws in dfs_by_sheet.items(): 
        base_df = ws[0] 
        print() 
        for i, df in enumerate(ws):
            diff = base_df.columns.difference(df.columns)
            if diff:
                print(f"{i:>2}. {ws_name[:12]} | {diff}")

############################
# 0, 0: 1
# 1, 0: 0
# 2, 0: 8
# 3, 0: 4
# 4, 9: 7
# 5, 0: 8
# 6, 0: 1
from combine_sheets import * 

file_list = [
    'corrected-28-sheets-cleaned.json',
]


file_idx, base_idx = 0, 0
wb_dfs = create_dfs(file_list[file_idx])

base_sheet_name = list(wb_dfs.keys())[base_idx]
base_wb = wb_dfs[base_sheet_name]
sheets = list(wb_dfs[base_sheet_name].keys()) 

dfs_sheets_compare = {}
for sheet in sheets:
    dfs_sheets_compare[sheet] = base_wb[sheet]

to_pop = []
for wb_name, wb in wb_dfs.items():
    counter = 0
    for sheet in sheets:
        df = wb[sheet]
        base_df = dfs_sheets_compare[sheet]
        in_base_not_in_other = base_df.columns.difference(df.columns)
        not_in_base_in_other = df.columns.difference(base_df.columns)

        if not in_base_not_in_other.empty or not not_in_base_in_other.empty:
            to_pop.append(wb_name)
            if counter == 0:
                print(f"\n\n\n\n=================\n\n{wb_name}")
                counter += 1
            print(f"\n\n---------\n{sheet}\n---------")
            print('\nIn Base:\n', in_base_not_in_other)
            print('\nNot In Base:\n', not_in_base_in_other)

to_pop = list(set(to_pop))
print(len(to_pop))
# print(to_pop)

for key in to_pop: 
    wb_dfs.pop(key)

print(wb_dfs.keys())


dfs_by_sheet = collect_dfs_by_sheet(wb_dfs)
sheet_dfs = join_dfs(dfs_by_sheet)
write_to_xlsx(sheet_dfs, file_list[file_idx])

# # Columns in df1 not in df2
# df1.columns.difference(df2.columns)

# # Columns in df2 not in df1
# df2.columns.difference(df1.columns)


wb_dfs_1 = make_dict(file_list[0])
wb_dfs_3 = make_dict(file_list[2])
wb_dfs_4 = make_dict(file_list[3])
wb_dfs_5 = make_dict(file_list[4])
wb_dfs_6 = make_dict(file_list[5])
wb_dfs_7 = make_dict(file_list[6])

wb_dfs_1_e = {k: v for k, v in wb_dfs_1.items() if k in excluded} 
wb_dfs_3_e = {k: v for k, v in wb_dfs_3.items() if k in excluded} 
wb_dfs_4_e = {k: v for k, v in wb_dfs_4.items() if k in excluded} 
wb_dfs_5_e = {k: v for k, v in wb_dfs_5.items() if k in excluded} 
wb_dfs_6_e = {k: v for k, v in wb_dfs_6.items() if k in excluded} 
wb_dfs_7_e = {k: v for k, v in wb_dfs_7.items() if k in excluded}

def make_dict(json_filename):
    print(f"Opening {json_filename}...")
    with open(json_filename, 'r') as f:
        a = f.read()
        sheets_dict = json.loads(a)

    return sheets_dict

excluded_dicts = {**wb_dfs_1_e, **wb_dfs_3_e, **wb_dfs_4_e, **wb_dfs_5_e, **wb_dfs_6_e, **wb_dfs_7_e}