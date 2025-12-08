
import openpyxl as xl
def info_snatcher(path: str,
                  work_sheet: str,
                  info_name, cat_row, search_row):
    wb_on_hand = xl.load_workbook(path)
    ws_on_hand = wb_on_hand[work_sheet]
    for index, cell in enumerate(ws_on_hand[cat_row]):
        if cell.value == info_name:
            target_index = index + 1 # 这是因为 enumerate 是从 0 开始计数，但 Excel 的列是从 1 开始计数，所以需要index + 1来转换。
    return ws_on_hand.cell(row = search_row, column = target_index).value


if __name__ == "__main__":
    test_path = "/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx"
    work_sheet = 'HVE 11_7 - J716C DVT-2 Build ru'
    test_info_name = "FSTP"
    test_cat_row = 2
    test_search_row = 6
    print(info_snatcher(test_path, work_sheet, test_info_name, test_cat_row, test_search_row))


