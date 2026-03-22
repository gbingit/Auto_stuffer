# 从Nimbus DVT2 BM上导出的excel中根据原则*拿出相应cfg的机器并计数
# *原则：1. 确定ADB CFG； 2. 在相应ADB CFG的CFG的Allocation as列找到对应的concatenator出来的值； 3. 拿出CFG，并且计数-1
# 在哪里（file_path)的什么sheet(work_sheet)中拿出多少个(quantity)什么cfg(cfg)的机器


# to improve the robustness of the code, better normalize every value when comparing

def out_taker(file_path: str,
              work_sheet: str,
              cfg: str,
              quantity: int ): ## 可以连续拿出来吗？（即不倒空现有plate）：应该不行，不然会覆盖盘子，上一个盘子就找不到了
    # import sys
    #
    from Normalization import to_normalize_cfg
    #
    # # 把dictionarization.py所在的目录路径添加到sys.path
    # sys.path.append(
    #     r"/Users/edison/PycharmProjects/PythonProject")  # 这里的r这里的 r 是 Python 里的原始字符串标记，作用是让字符串里的特殊字符（比如\、空格等）“原样保留”，不会被 Python 解析成转义符。# 举个例子：如果直接写路径 "C:\Users\test"，Python 会把 \U 当成转义符，导致报错；但加了 r 写成 r"C:\Users\test"，Python 就会把整个字符串当成 “原始内容”，不会解析转义，路径就能正确识别。
    from dictionarization import to_dic_head_col
    from openpyxl.styles import PatternFill
    import openpyxl as xl
    wb = xl.load_workbook(file_path)
    ws = wb[work_sheet]
    head_col_dic = to_dic_head_col(ws) # （这里可将function改良成可输入参数第几行是head，默认是1)
    # print(head_col_dic)
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    plate = {} # 先来个空盘子
    for row in range(2, ws.max_row + 1):
        target_row = None
        if to_normalize_cfg(ws.cell(row, head_col_dic['Name']).value) == to_normalize_cfg(cfg): # 在Name列找目标cfg所在列
            target_row = row # 定为目标行
            print(target_row)
            # print(type(ws.cell(target_row, head_col_dic['Input Qty']).value))
            break
    if not target_row:
        return 'No cfg found'
    elif int(ws.cell(target_row, head_col_dic['Input Qty']).value) >= int(quantity): # 如果目标cfg数量大于要拿的量，就开始拿出来（减去拿出来的数量）并标记红色
        ws.cell(target_row, head_col_dic['Input Qty']).value -= quantity
        ws.cell(target_row, head_col_dic['Input Qty']).fill = red_fill
        wb.save(file_path)
    else:
        return f'No enough cfgs to take, leftover is {ws.cell(target_row, head_col_dic['Input Qty']).value}'

    plate[cfg] = quantity # 只有当找到目标cfg并且所剩数量充足的情况下才会走到这里： 所以就开始装盘
    return plate

# 测试：从/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/J716 DVT-2 FATP (C) 2. QSMC Config (21022).xlsx 中拿5台DVT2c-D7-EMC4-I-N-128M
if __name__ == '__main__':
    file_path = '/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/J716 DVT-2 FATP (C) 2. QSMC Config (21022).xlsx'
    work_sheet = 'FATP (C)'
    cfg = 'DVT2c-D16-A-B-128S'
    quantity = 5
    print(out_taker(file_path,
              work_sheet,
              cfg,
              quantity))




