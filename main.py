#1. HVE
# 1.1. 先从BR里将对应列的一位数字编码（在这里是B列第3行到第14行）组合到HVE后，得到HVE1-HVE12；#
import openpyxl as xl

from Normalization import to_normalize, to_lsplit_meaning
from cell_value_concatenator import cell_value_concatenator
from dictionarization import to_dic_head_col
from to_set_order import to_set_order
S2F_file_path = '/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/练手 DVT-2 FATP (C).xlsx'
config_file_path = '/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/J716 DVT-2 FATP (C) 2. QSMC Config (21022).xlsx'
BR_file_path = "/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx"
HVE_work_sheet = 'HVE 11_7 - J716C DVT-2 Build ru'
S2F_file = xl.load_workbook(filename=S2F_file_path)
config_file = xl.load_workbook(config_file_path)
BR_file = xl.load_workbook(BR_file_path)
S2F_sheet = S2F_file['Shipments']
config_sheet = config_file.active
config_sheet_head = to_dic_head_col(config_sheet)

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





# 1.4. 再从1.3.中确定的config行里out-take出一个盘子里装着{"DVT2c-Mini1-A-N-36M":8}，同时就在Config表中的同行“Input Qty”列减去要拿的数量，如果不够就显示不够
## 是一把拿到位，还是一个一个拿？


# 1.5. 再把从1.4.拿出的盘子里的cfg一个一个放到STF的相应位置里





