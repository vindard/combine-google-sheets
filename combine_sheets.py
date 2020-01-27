import pandas as pd
import json
with open('24-sheets.json', 'r') as f:
    a = f.read()
    sheets_dict = json.loads(a)

wb_dfs = {}
# wb1 = a['Copy of Suthahgar - 04100-70 - Household Survey Data Entry ']
for wb_name, wb in sheets_dict.items():
    wb_dfs[wb_name] = {}
    # ws1_1 = wb1['Cover Sheets']
    for ws_name, ws in wb.items():
        ws_headers =  ws['headers']
        ws_vals = ws['vals']
        ws_tuples = [tuple(i) for i in  ws_headers]
        headers = pd.MultiIndex.from_tuples(ws_tuples)
        df = pd.DataFrame(ws_vals, columns=headers)

        wb_dfs[wb_name][ws_name] = df


# Get all dfs per wb
dfs_by_sheet = {}
for wb_name, wb_df in wb_dfs.items():
    for ws_name, ws_df in wb_df.items():
        if not dfs_by_sheet.get(ws_name):
            dfs_by_sheet[ws_name] = []
        dfs_by_sheet[ws_name].append(ws_df)

sheet_dfs = {}
for ws_name, ws_df_list in dfs_by_sheet.items():
    sheet_dfs[ws_name] = pd.concat(ws_df_list, ignore_index=True)

with pd.ExcelWriter('output.xlsx') as writer:  # doctest: +SKIP
    for ws_name, df in sheet_dfs.items():
        df.to_excel(writer, sheet_name=ws_name)

