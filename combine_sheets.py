import pandas as pd
import json

def create_dfs(json_filename):
    print(f"Opening {json_filename}...")
    with open(json_filename, 'r') as f:
        a = f.read()
        sheets_dict = json.loads(a)

    print(f"Making all dataframes...")
    wb_dfs = {}
    for wb_name, wb in sheets_dict.items():
        wb_dfs[wb_name] = {}
        for ws_name, ws in wb.items():
            ws_headers =  ws['headers']
            ws_vals = ws['vals']
            ws_tuples = [tuple(i) for i in  ws_headers]
            headers = pd.MultiIndex.from_tuples(ws_tuples)
            df = pd.DataFrame(ws_vals, columns=headers)

            wb_dfs[wb_name][ws_name] = df

    return wb_dfs

def collect_dfs_by_sheet(wb_dfs):
    # Get all dfs per wb
    print(f"Rearranging all dataframes by sheet...")
    dfs_by_sheet = {}
    for wb_name, wb_df in wb_dfs.items():
        for ws_name, ws_df in wb_df.items():
            if not dfs_by_sheet.get(ws_name):
                dfs_by_sheet[ws_name] = []
            dfs_by_sheet[ws_name].append(ws_df)

    print(f"Returning all dataframes by sheet.")
    return dfs_by_sheet


def join_dfs(dfs_by_sheet):
    print(f"Joining wb sheet dataframes into single dataframe...")
    sheet_dfs = {}
    for ws_name, ws_df_list in dfs_by_sheet.items():
        # if not 'B3. - Animals Likely To Be Owned' in ws_name:
        #     print('HERE', ws_name)
        sheet_dfs[ws_name] = pd.concat(ws_df_list, ignore_index=True)

    return sheet_dfs

def write_to_xlsx(sheet_dfs, json_filename):
    output_file = f"{json_filename[:-5]}.xlsx"
    print(f"Writing resultx to {output_file}")
    with pd.ExcelWriter(output_file) as writer:
        for ws_name, df in sheet_dfs.items():
            df.to_excel(writer, sheet_name=ws_name)


def make_xlsx(json_filename):
    wb_dfs = create_dfs(json_filename)
    dfs_by_sheet = collect_dfs_by_sheet(wb_dfs)
    sheet_dfs = join_dfs(dfs_by_sheet)
    write_to_xlsx(sheet_dfs, json_filename)

    return sheet_dfs
