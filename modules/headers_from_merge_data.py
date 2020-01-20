from libs.config import open_workbook
from libs.utils import make_headers

# NOTE: Method comes from `gspread-pandas package. This method
#       appears to currently be broken. Test against my own
#       method in libs.headers_from_cell_parse to fix.

def _fetch_merges(workbook_name, sheet_num=0, client=None):
    try:
        wb = open_workbook(workbook_name, client)
    except:
        return []

    meta = wb.fetch_sheet_metadata()
    ws_meta = meta['sheets'][sheet_num]
    merges = ws_meta['merges']

    return merges

def _fix_merge_values(all_vals, workbook_name='', sheet_num=0, client=None):
    """
    Assign the top-left value to all cells in a merged range.

    Parameters
    ----------
    all_vals : list
        Values returned by
        :meth:`get_all_values() <gspread.models.Sheet.get_all_values()>_`


    Returns
    -------
    list
        Fixed values
    """
    merges = _fetch_merges(workbook_name, sheet_num)

    for merge in merges:
        start_row, end_row = merge["startRowIndex"], merge["endRowIndex"]
        start_col, end_col = (merge["startColumnIndex"], merge["endColumnIndex"])

        # ignore merge cells outside the data range
        if start_row < len(all_vals) and start_col < len(all_vals[0]):
            orig_val = all_vals[start_row][start_col]
            for row in all_vals[start_row:end_row]:
                row[start_col:end_col] = [
                    orig_val for i in range(start_col, end_col)
                ]

    return all_vals

def separate_headers_vals(all_vals, starts_at=1, num_rows=1, workbook_name='', sheet_num=0):
    all_vals = _fix_merge_values(all_vals, workbook_name, sheet_num)
    headers, vals = make_headers(all_vals, starts_at, num_rows)

    return headers, vals