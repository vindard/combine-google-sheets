from libs.config import MASTER_WORKBOOK
from modules.build_sheet_dicts import run as build_sheet_dicts

def fetch_sheet_headers(sheet_dicts, wb_name, sheet_num):
    sheet_idx = sheet_num - 1
    keys = list(sheet_dicts.keys())

    # Fetch ws1 dict for sheet_num
    wb = sheet_dicts[wb_name]
    wb_names = list(wb.keys())

    wb_compare_sheet_name = wb_names[sheet_idx]
    wb_compare_sheet = wb[wb_compare_sheet_name]
    wb_sheet_headers = wb_compare_sheet['headers']

    return wb_sheet_headers

def run_compare(sheet_dicts, wb_idx=0):
    assert len(sheet_dicts) >= wb_idx + 1, \
        f"\"wb_idx\" ('{wb_idx}') is bigger than length of sheet_dicts ('{len(sheet_dicts)}')"

    compare_report, report_body = '', ''
    keys = list(sheet_dicts.keys())
    base_wb_name = MASTER_WORKBOOK or keys[0]
    base_sheets_names = list(sheet_dicts[base_wb_name].keys())
    compare_wb_name = keys[wb_idx]

    base_headers = {sheet_name: fetch_sheet_headers(sheet_dicts, base_wb_name, sheet_num=i)
                        for i, sheet_name in enumerate(base_sheets_names)}

    compare_report += "\n====================\n\n"
    compare_report += f"Running compare for workbooks:\n- '{base_wb_name}'\n- '{compare_wb_name}'\n"
    for i, sheet_name in enumerate(base_sheets_names):
        h1 = base_headers[sheet_name]
        h2 = fetch_sheet_headers(sheet_dicts, compare_wb_name, sheet_num=i)

        if h1 == h2:
            continue

        report_body += f"\n\nSheet #{i+1}: {sheet_name}\n"
        report_body += f"---\n"
        hs1, hs2 = set(h1), set(h2)
        set_diff = hs1 - hs2
        if set_diff:
            for i, val in enumerate(set_diff):
                if val in h1:
                    report_body += f"{i+1}. Missing in compare sheet:\n  {val}\n"
                else:
                    report_body += f"{i+1}. Extra in compare sheet (not in base sheet):\n  {val}\n"
        else:
            report_body += f"Headers don't match against base sheet, please review.\n"

    compare_report += report_body or "\nNo differences found!"

    return bool(report_body), compare_report


if __name__ == '__main__':
    sheet_dicts = build_sheet_dicts()
    diff, results = run_compare(sheet_dicts, 0)
    if diff:
        print(results)
