import pandas as pd
import json

def create_dicts(json_filename):
    print(f"Opening {json_filename}...")
    with open(json_filename, 'r') as f:
        wbs = f.read()
        sheets_dict = json.loads(wbs)

    return sheets_dict



def clean_wildlife(json_filename='00-24-sheets.json'):
    wbs = create_dicts(json_filename)

    for wb_name, wb in wbs.items():
        ws_name = 'A. Wildlife Experiences'
        ws = wb[ws_name]

        vals = []
        for val in ws['vals']:
            check = ''.join(val) != "0"
            if check:
                vals.append(val)

        wbs[wb_name][ws_name]['vals'] = vals

        print()
        print(wb_name)
        print(len(ws['vals']), '|', len(vals))

    return wbs
