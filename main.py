#1. HVE
# 1.1. 先从BR里将对应列的一位数字编码（在这里是B列第3行到第14行）组合到HVE后，得到HVE1-HVE12；#
import sys
import os
from zoneinfo import available_timezones

# 获取当前脚本所在目录的绝对路径
current_dir = os.path.dirname(os.path.abspath(__file__))
print(current_dir)
# 把当前目录加入模块搜索路径
sys.path.append(current_dir)

import openpyxl as xl

from Normalization import to_normalize, to_lsplit_meaning, to_normalize_cfg
from cell_value_concatenator import cell_value_concatenator
from dictionarization import to_dic_head_col
from to_set_order import to_set_order
from out_taker import out_taker

S2F_file_path = '/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/练手 DVT-2 FATP (C).xlsx'
config_file_path = '/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/J716 DVT-2 FATP (C) 2. QSMC Config (21022).xlsx'
BR_file_path = "/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx"
HVE_work_sheet = 'HVE 11_7 - J716C DVT-2 Build ru'
S2F_file = xl.load_workbook(filename=S2F_file_path)
config_file = xl.load_workbook(config_file_path)
BR_file = xl.load_workbook(BR_file_path)
S2F_sheet = S2F_file['Shipments']
config_sheet = config_file['FATP (C)']
BR_HVE_sheet = BR_file['HVE 11_7 - J716C DVT-2 Build ru']
config_sheet_head = to_dic_head_col(config_sheet)
S2F_head_dic = to_dic_head_col(S2F_sheet)
BR_HVE_head_dic = to_dic_head_col(BR_HVE_sheet, head_row= 2)
# print(BR_HVE_head_dic)

for row in range(3, 15):
    target_HVE_code = cell_value_concatenator("/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx",
                            'HVE 11_7 - J716C DVT-2 Build ru', cell1 = 'D18', cell2 = f"B{row}"
                            )

# 1.2. 再加上BR里Qty那一列的数量合成一个order dic，e,g,.{"HVE1":8}; #

    HVE_code_order = to_set_order("/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx",
                            'HVE 11_7 - J716C DVT-2 Build ru', row, 15, target_HVE_code)
    print(HVE_code_order)
# 1.3. 再将得到的HVE1-HVE12和Nimbus上的Config表中的Allocations as列里的内容做对比，找到相同的定位到同一行的Name列的CFG；>
    target_cfgs = []
    for row2 in range(1, config_sheet.max_row + 1):
        if to_normalize(target_HVE_code) in to_lsplit_meaning(config_sheet.cell(row2, config_sheet_head['Allocations as']).value): #如果有两个或以上符合条件的cfg name怎么办？？(自己选) ✅ ##这里如果找HVE1，那么HVE10，11，12都会被搜寻出来，该怎么处理？（Python把某个单元格的值按照遇到“,” 和空格就分开的原则split，再把split的结果存到一个列表当中，然后再和列表中的单个元素做整体匹配（在Normalization中定义新的功能））✅
            target_cfgs.append(config_sheet.cell(row2, config_sheet_head['Name']).value)
    if len(target_cfgs) > 1:
        print(f"there are multiple cfgs designated to {target_HVE_code}, they are {target_cfgs}")
        target_cfg = input(f"which cfg do you like to allocate to {target_HVE_code}? ")
    elif len(target_cfgs) == 1:
        target_cfg = target_cfgs[0]
    else:
        print(f"there are no cfgs designated to {target_HVE_code}")
    if target_cfg:
        cfg_order = to_set_order("/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx",
                            'HVE 11_7 - J716C DVT-2 Build ru', row, 15, target_cfg)
        HVE_code_order[target_HVE_code] = cfg_order
        print(HVE_code_order)
        # print(type(cfg_order[target_cfg]))





# 1.4. 再从1.3.中确定的config行里out-take出一个盘子里装着{"DVT2c-Mini1-A-N-36M":8}，同时就在Config表中的同行“Input Qty”列减去要拿的数量，如果不够就显示不够
## 是一把拿到位，还是一个一个拿？
        plate = out_taker(config_file_path, 'FATP (C)', target_cfg, cfg_order[target_cfg])
        print(f'已经从Build Matrix中拿出{cfg_order[target_cfg]}个{target_cfg}，所以现在的盘子里装着{cfg_order[target_cfg]}个{target_cfg}')



# 1.5. 再把从1.4.拿出的盘子里的cfg一个一个放到STF的相应位置里
## 如果S2F里location那一列的ADB cfg = BR 里的 No那一列，且S2F’Config 是空的，就定义一个空位，然后从盘子里拿出来一个放进去，直到盘子放完或者没有空位了
        if type(plate) == dict:
            for rows_S2F in range(1, S2F_sheet.max_row + 1):
                    if plate[target_cfg] == 0:
                        break
                    elif to_normalize_cfg(S2F_sheet.cell(rows_S2F, S2F_head_dic['Location']).value) \
                    == to_normalize_cfg(BR_HVE_sheet.cell(row, BR_HVE_head_dic['Config']).value) \
                    and not S2F_sheet.cell(rows_S2F, S2F_head_dic['Config']).value:

                        print(rows_S2F)
                        available_cell = S2F_sheet.cell(rows_S2F, S2F_head_dic['Config'])
                        plate[target_cfg] -= 1
                        print(plate[target_cfg])
                        available_cell.value = target_cfg
            S2F_file.save(S2F_file_path)

        else:
            input(f'{plate}, what should be done? ')








