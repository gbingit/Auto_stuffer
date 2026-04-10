def to_dic_head_index(sheet, head_row: int = 1): #输出表头的index字典（从0开始）
    index_lib_0 = {}
    for index, cell in enumerate(sheet[head_row]):
        if cell.value is not None and cell.value != '':
            index_lib_0[cell.value] = index
    return index_lib_0


def to_dic_head_col(sheet, head_row: int = 1): # 输出表头的列数字典（从1开始）
    col_count_lib_0 = {}
    for index, cell in enumerate(sheet[head_row]):
        if cell.value is not None and cell.value != '':
            col_count_lib_0[cell.value] = index + 1
    return col_count_lib_0



