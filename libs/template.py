from libs.config import open_workbook, MASTER_WORKBOOK

def generate_json_template(wb_name: str = MASTER_WORKBOOK) -> dict:
    name, wb = open_workbook(wb_name)
    sheet_dict = {}
    for i, sheet in enumerate(wb):
        entries_dict = {}
        entries_dict['sheet'] = "sheet-object-goes-here"
        entries_dict['sheet_idx'] = i
        entries_dict['header_info_dict'] = {"start_at": 1, "num_rows": 1}

        sheet_dict[sheet.title] = entries_dict

    return sheet_dict
