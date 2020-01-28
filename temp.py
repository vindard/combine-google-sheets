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

############################
# 2, 0: 8
# 3, 0: 4
# 4, 9: 7
# 5, 0: 8
# 6, 0: 1

from combine_sheets import * 

file_idx, base_idx = 4, 9
wb_dfs = create_dfs(file_list[file_idx])

base_sheet_name = list(wb_dfs.keys())[base_idx]
base_wb = wb_dfs[base_sheet_name]
sheets = list(wb_dfs[base_sheet_name].keys()) 

dfs_sheets_compare = {}
for sheet in sheets:
    dfs_sheets_compare[sheet] = base_wb[sheet]

to_pop = []
for wb_name, wb in wb_dfs.items():
    for sheet in sheets:
        df = wb[sheet]
        base_df = dfs_sheets_compare[sheet]

        in_base_not_in_other = base_df.columns.difference(df.columns)
        not_in_base_in_other = df.columns.difference(base_df.columns)

        if not in_base_not_in_other.empty or not not_in_base_in_other.empty:
            to_pop.append(wb_name)

to_pop = list(set(to_pop))
print(len(to_pop))

for key in to_pop: 
    wb_dfs.pop(key)

dfs_by_sheet = collect_dfs_by_sheet(wb_dfs)
join_and_output_xlsx(dfs_by_sheet, file_list[file_idx])
        
# # Columns in df1 not in df2
# df1.columns.difference(df2.columns)

# # Columns in df2 not in df1
# df2.columns.difference(df1.columns)

