#1. HVE
# 1.1. 先从BR里将对应列的一位数字编码（在这里是B列第3行到第14行）组合到HVE后，得到HVE1-HVE12；#
import sys
import os
import configparser
from zoneinfo import available_timezones

# 获取当前脚本所在目录的绝对路径
current_dir = os.path.dirname(os.path.abspath(__file__))
print(current_dir)
# 把当前目录加入模块搜索路径
sys.path.append(current_dir)
# try
import openpyxl as xl

from Normalization import to_normalize, to_lsplit_meaning, to_normalize_cfg
from cell_value_concatenator import cell_value_concatenator
from dictionarization import to_dic_head_col
from to_set_order import to_set_order
from out_taker import out_taker


# ---------------------- 新增：读取配置文件 ----------------------
def load_config(config_path="config.ini"):
    config = configparser.ConfigParser()
    if not os.path.exists(config_path):
        print(f"❌ 配置文件 {config_path} 不存在，请检查路径！")
        sys.exit(1)
    config.read(config_path, encoding="utf-8")
    return config

config = load_config()

# 从配置文件获取参数（用户只需改config.ini，不用动这里）
S2F_file_path = config["FILES"]["S2F_file_path"]
config_file_path = config["FILES"]["config_file_path"]
BR_file_path = config["FILES"]["BR_file_path"]
HVE_work_sheet = config["SHEETS"]["HVE_work_sheet"]
S2F_sheet_name = config["SHEETS"]["S2F_sheet_name"]
config_sheet_name = config["SHEETS"]["config_sheet_name"]
BR_HVE_sheet_name = config["SHEETS"]["BR_HVE_sheet_name"]
BR_start_row = int(config["RANGE"]["BR_start_row"])
BR_end_row = int(config["RANGE"]["BR_end_row"])

# S2F_file_path = '/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/练手 DVT-2 FATP (C).xlsx'
# config_file_path = '/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/J716 DVT-2 FATP (C) 2. QSMC Config (21022).xlsx'
# BR_file_path = "/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx"
# HVE_work_sheet = 'HVE 11_7 - J716C DVT-2 Build ru'

def main():
    """核心执行函数：封装所有逻辑，用户只需运行这个函数"""
    print("🔄 Staring to stuff...")
    input("Press enter to continue...")
    try:
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
            print(f"\n📌 processing row No.{row} in BR...")
            input("Press enter to continue...")
            target_HVE_code = cell_value_concatenator("/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx",
                                    'HVE 11_7 - J716C DVT-2 Build ru', cell1 = 'D18', cell2 = f"B{row}"
                                    )
            print(f"✅ HVEcode, {target_HVE_code}, generated")
            input("Press enter to continue...")


        # 1.2. 再加上BR里Qty那一列的数量合成一个order dic，e,g,.{"HVE1":8}; #

            HVE_code_order = to_set_order("/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx",
                                    'HVE 11_7 - J716C DVT-2 Build ru', row, 15, target_HVE_code)
            print(f"✅ HVE_code_"
                  f"order, {HVE_code_order}, generated")
            input("Press enter to continue...")
        # 1.3. 再将得到的HVE1-HVE12和Nimbus上的Config表中的Allocations as列里的内容做对比，找到相同的定位到同一行的Name列的CFG；>
            target_cfgs = []
            for row2 in range(1, config_sheet.max_row + 1):
                if to_normalize(target_HVE_code) in to_lsplit_meaning(config_sheet.cell(row2, config_sheet_head['Allocations as']).value): #如果有两个或以上符合条件的cfg name怎么办？？(自己选) ✅ ##这里如果找HVE1，那么HVE10，11，12都会被搜寻出来，该怎么处理？（Python把某个单元格的值按照遇到“,” 和空格就分开的原则split，再把split的结果存到一个列表当中，然后再和列表中的单个元素做整体匹配（在Normalization中定义新的功能））✅
                    target_cfgs.append(config_sheet.cell(row2, config_sheet_head['Name']).value)
            if len(target_cfgs) > 1:
                print(f"there are multiple cfgs designated to {target_HVE_code}, they are {target_cfgs}")
                target_cfg = input(f"which cfg do you like to allocate to {target_HVE_code}? ")
                # 增加输入校验
                while target_cfg not in target_cfgs:
                    print(f"❌ Wrong input! please choose from the list {target_cfgs}")
                    target_cfg = input("please input again：")
            elif len(target_cfgs) == 1:
                target_cfg = target_cfgs[0]
                print(f"✅ matched only with {target_cfg}")
                input("Press enter to continue...")
            else:
                print(f"there are no cfgs designated to {target_HVE_code}, please coordinate with DRI on solution")
                input("Press enter to continue...")
                continue ## 可以直接跳出循环，再继续下一个row（很重要）！不然即使没找到target_cfg也会继续跑后面的装盘扣库存等等，浪费算力！！
            if target_cfg:
                cfg_order = to_set_order("/Users/edison/Desktop/J716 S2F/DVT2/FATP/C/stuff 练手/PWZM DVT2c Build Rules V5.xlsx",
                                    'HVE 11_7 - J716C DVT-2 Build ru', row, 15, target_cfg)
                HVE_code_order[target_HVE_code] = cfg_order
                print(f'HVE_code_order updated to {HVE_code_order}')
                print(f'{target_HVE_code} shoule be having {cfg_order[target_cfg]} {target_cfg}\n')
                input("Press enter to continue...")
                # print(type(cfg_order[target_cfg]))





        # 1.4. 再从1.3.中确定的config行里out-take出一个盘子里装着{"DVT2c-Mini1-A-N-36M":8}，同时就在Config表中的同行“Input Qty”列减去要拿的数量，如果不够就显示不够
        ## 是一把拿到位，还是一个一个拿？
                plate = out_taker(config_file_path, 'FATP (C)', target_cfg, cfg_order[target_cfg])
                # （如果BM中数量不够怎么办？ 会没有plate
                if isinstance(plate, str):  # out_taker返回错误信息
                    print(f"❌ {plate}") # （这里可能是没有找到cfg，也可能是找到了但是数量不够
                    print("please pause and ponder how to deal with it")
                    input("Press enter to continue...")
                    continue
                print(f'已经从Build Matrix中拿出{cfg_order[target_cfg]}个{target_cfg} and the plate looks like {plate}')





        # 1.5. 再把从1.4.拿出的盘子里的cfg一个一个放到STF的相应位置里
        ## 如果S2F里location那一列的ADB cfg = BR 里的 Config那一列，且S2F’Config 是空的，就定义一个空位，然后从盘子里拿出来一个放进去，直到盘子放完或者没有空位了
                filled_count = 0
                if type(plate) == dict:
                    print(f'Now fill in {target_HVE_code} cfg...\n')
                    for rows_S2F in range(1, S2F_sheet.max_row + 1):

                        if plate[target_cfg] == 0:
                            break
                        elif to_normalize_cfg(S2F_sheet.cell(rows_S2F, S2F_head_dic['Location']).value) \
                        == to_normalize_cfg(BR_HVE_sheet.cell(row, BR_HVE_head_dic['Config']).value) \
                        and not S2F_sheet.cell(rows_S2F, S2F_head_dic['Config']).value:

                            # print(f'In S2F row No.{rows_S2F} needs {S2F_sheet.cell(rows_S2F, S2F_head_dic['Location']).value}; In BR row No. {row} requires {BR_HVE_sheet.cell(row, BR_HVE_head_dic['Config']).value} also; so S2F row No.{rows_S2F} is the target row to fill in the cfg {target_cfg} ')
                            available_cell = S2F_sheet.cell(rows_S2F, S2F_head_dic['Config'])
                            # print(f'--> Now take 1 {target_cfg} from the plate')
                            plate[target_cfg] -= 1 #先从盘里拿出来
                            # print(f'--> Now the plate looks like {plate}') #再看看盘里还剩多少
                            # print(f'--> Now put the {target_cfg} in the S2F row No.{rows_S2F}')
                            available_cell.value = target_cfg #再放进available的cell里面
                            # print(f'--> Now the S2F row No.{rows_S2F} has {target_cfg} filled in\n\n')
                            filled_count += 1
                    S2F_file.save(S2F_file_path)
                    print(f"\n✅  {filled_count} cells have been filled，and doc saved")
                    input("Press enter to continue...")
        print("🎉 All HVE allocations have been stuffed！")




    except Exception as e:
        # 错误捕获，给出友好提示
        print(f"\n❌ 执行出错：{str(e)}")
        # 可选：保存错误日志到文件
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{os.path.basename(__file__)} 执行错误：{str(e)}\n")
        input("按回车键退出...")
        sys.exit(1)

# ---------------------- 入口函数 ----------------------
if __name__ == "__main__":
    # 欢迎语
    print("="*50)
    print("🎯 机器自动分配工具 V1.0")
    print("="*50)
    input("按回车键开始执行（请确保已配置好config.ini文件）...")
    main()
    input("\n✅ 执行完成，按回车键退出...")






