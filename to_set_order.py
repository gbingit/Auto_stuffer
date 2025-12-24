# from dictionarization import to_dic_head_col
from dictionarization import to_dic_head_col


def to_set_order(file_path, work_sheet, row, col, name):
    import openpyxl as xl
    wb = xl.load_workbook(file_path)
    ws = wb[work_sheet]
    # to_dic_head_col(ws)
    qty = ws.cell(row, col).value
    plate = {}
    plate[name] = qty
    return plate

if __name__ == '__main__':
    file_path = "/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx"
    work_sheet = 'HVE 11_7 - J716C DVT-2 Build ru'
    row = 3
    col = 15
    print(to_set_order(file_path, work_sheet, row, col, 'HVE1'))


