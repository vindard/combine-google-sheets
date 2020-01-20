from libs.utils import make_headers

def _fix_merge_headers(merged_headers):
    if not merged_headers:
        return []

    unmerged_headers = []
    curr = [''] * len(merged_headers[0])
    for col in merged_headers:
        new_col = []
        for i, item in enumerate(col):
            if i == 0 and item:
                curr = list(col)
            new_col.append(item or curr[i])
            curr[i] = item or curr[i]
            # if len(new_col) > 1 and not new_col[-1]:
            #     new_col[-1] = new_col[-2]
        unmerged_headers.append(new_col)

    # Fill in blanks cells
    for i, header_col in enumerate(unmerged_headers):
        for j, cell in enumerate(header_col):
            if j != 0 and cell == '':
                unmerged_headers[i][j] = unmerged_headers[i][j-1]

    # Fix blank top rows
    m = 0
    for m, val in enumerate(unmerged_headers):
        if val != "":
            break
    for n in range(m):
        unmerged_headers[n] = val

    return [tuple(header_list) for header_list in unmerged_headers]


def separate_headers_vals(all_vals, starts_at=1, num_rows=1):
    headers, vals = make_headers(all_vals, starts_at, num_rows)

    return _fix_merge_headers(headers), vals
