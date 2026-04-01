import sys
import os
import configparser
import openpyxl as xl
from tkinter import Tk, filedialog, messagebox

def info_snatcher(path: str,
                  work_sheet: str,
                  info_name: str,
                  cat_row: str,
                  search_row: str):
    wb_on_hand = xl.load_workbook(path)
    ws_on_hand = wb_on_hand[work_sheet]
    for index, cell in enumerate(ws_on_hand[cat_row]):
        if cell.value == info_name:
            target_index = index + 1 # 这是因为 enumerate 是从 0 开始计数，但 Excel 的列是从 1 开始计数，所以需要index + 1来转换。
    return ws_on_hand.cell(row = search_row, column = target_index).value

def cell_value_concatenator(file_path: str,
                            work_sheet: str,
                            cell1: str,
                            cell2: str,
                            output_cell: str = ''):
    wb = xl.load_workbook(file_path)
    ws = wb[work_sheet]
    val1 = str(ws[cell1].value)
    val2 = str(ws[cell2].value)
    if output_cell:
        ws[output_cell].value = val1 + val2
    return val1 + val2

def to_set_order(file_path, work_sheet, row, col, name):
    import openpyxl as xl
    wb = xl.load_workbook(file_path)
    ws = wb[work_sheet]
    # to_dic_head_col(ws)
    qty = ws.cell(row, col).value
    plate = {}
    plate[name] = qty
    return plate

def out_taker(file_path: str,
              work_sheet: str,
              cfg: str,
              quantity: int ): ## 可以连续拿出来吗？（即不倒空现有plate）：应该不行，不然会覆盖盘子，上一个盘子就找不到了
    # import sys
    # # 把dictionarization.py所在的目录路径添加到sys.path
    # sys.path.append(
    #     r"/Users/edison/PycharmProjects/PythonProject")  # 这里的r这里的 r 是 Python 里的原始字符串标记，作用是让字符串里的特殊字符（比如\、空格等）“原样保留”，不会被 Python 解析成转义符。# 举个例子：如果直接写路径 "C:\Users\test"，Python 会把 \U 当成转义符，导致报错；但加了 r 写成 r"C:\Users\test"，Python 就会把整个字符串当成 “原始内容”，不会解析转义，路径就能正确识别。
    from openpyxl.styles import PatternFill
    wb = xl.load_workbook(file_path)
    ws = wb[work_sheet]
    head_col_dic = to_dic_head_col(ws) # （这里可将function改良成可输入参数第几行是head，默认是1)
    # print(head_col_dic)
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    plate = {} # 先来个空盘子
    for row in range(2, ws.max_row + 1):
        target_row = None
        if to_normalize_cfg(ws.cell(row, head_col_dic['Name']).value) == to_normalize_cfg(cfg): # 在Name列找目标cfg所在列
            target_row = row
            print(target_row)
            break
    if not target_row:
        return 'No cfg found'
    elif int(ws.cell(target_row, head_col_dic['Input Qty']).value) >= int(quantity): # 如果目标cfg数量大于要拿的量，就开始拿出来（减去拿出来的数量）并标记红色
        ws.cell(target_row, head_col_dic['Input Qty']).value -= quantity
        ws.cell(target_row, head_col_dic['Input Qty']).fill = red_fill
        wb.save(file_path)
    else:
        return f'No enough cfgs to take, leftover is {ws.cell(target_row, head_col_dic['Input Qty']).value}'

    plate[cfg] = quantity
    return plate

def to_normalize(text):
    return str(text).lower().replace(' ', '').replace(',', '')

def to_normalize_address(address):
    return str(address).lower().replace(' ','').replace(',', '').replace('Attn:', '')

def to_normalize_cfg(cfg):
    return str(cfg).lower().replace(' ', '').replace(',', '').replace('-', '')

def to_lsplit_meaning(text):
    tem_body = str(text).lower().replace(',', ' ').replace('*', ' ')
    return tem_body.split()

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

# ====================== 【小白核心封装】 ======================
# 自动隐藏命令行黑框（Windows）
try:
    import ctypes
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 1)
except:
    pass

# 自动获取当前文件夹
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(CURRENT_DIR)
#保证找到依赖文件, such as out_taker.py, to_set_order.py, etc.

# ====================== 配置文件自动管理 ======================
CONFIG_PATH = os.path.join(CURRENT_DIR, "自动配置.ini")
#先在当前文件夹地址自动生成一个路径到"自动配置.ini"，比如当前文件夹地址a/b/c，那么就自动化生产a/b/c/自动配置.ini,先占个位

def load_or_create_config(): #总结：创建或读取（已存在的话）"自动配置.ini"，返回一个已经读取过的对象（config）以供调用init里面的内容
    """自动创建/读取配置文件，小白完全不用管"""
    config = configparser.ConfigParser()
    #用来读写"自动配置.ini"的操作工具，有了它才能读 .ini 文件；写 .ini 文件；按 [分组] 读取内容；存键值对：key=value

    if not os.path.exists(CONFIG_PATH):
        # 首次运行：自动生成配置
        config["FILES"] = {
            "S2F_file_path": "",
            "config_file_path": "",
            "BR_file_path": ""
        }
        config["SHEETS"] = {
            "HVE_work_sheet": "HVE 11_7 - J716C DVT-2 Build ru",
            "S2F_sheet_name": "Shipments",
            "config_sheet_name": "FATP (C)",
            "BR_HVE_sheet_name": "HVE 11_7 - J716C DVT-2 Build ru"
        }
        config["RANGE"] = {
            "BR_start_row": "3",
            "BR_end_row": "14"
        }
        #这里只是先在内存里写好稍后要写进"自动配置.ini"里的内容，但是房子还没盖好东西进不去
        #这里的config是个什么数据结构? 是个ConfigParser，但实质就是嵌套字典，调用的时候用config["FILES"]["S2F_file_path"]这样双重键

        with open(CONFIG_PATH, "w", encoding="utf-8") as f: #这里是盖房子的地方，open方法发现没有的话会自动创建一个空白的 自动配置.ini
            config.write(f) #并且写进去东西
        print("✅ 首次运行：已自动生成配置文件")

    config.read(CONFIG_PATH, encoding="utf-8")
    return config #这里return的到底是啥？**load_or_create_config () 返回：一个已经加载好所有配置的 config 对象
# 程序后面所有路径、表名、行号，全部从这个返回值里拿！e.g., cfg = config["FILES"]
#     sheet_cfg = config["SHEETS"]
#     range_cfg = config["RANGE"]

# ====================== 图形化选择Excel文件（小白点一下） ======================
def select_excel_file(title="请选择Excel文件"):
    """弹出窗口选文件，完全不用写路径""" #窗口的实质就是一块内存，可以允许用户输入而已
    root = Tk() #先创建一个空窗口（tkinter的规则，必须要有一个父窗口）
    root.withdraw() #直接隐藏，因为后面用不到这个空父窗口
    root.attributes('-topmost', True)  #让这个窗口一直处于最前面，保证不会被遮挡
    file_path = filedialog.askopenfilename( #在弹窗里选择一个文件+选好后返回文件绝对路径
        title=title, #窗口标题
        filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")], #默认只能选xlsx文件，但如果非要选其他也能在下拉格式里选择（后路）
        initialdir=CURRENT_DIR #初始在当前文件夹地址里
    )
    return file_path if file_path else None

# ====================== 统一消息提示 ======================
def show_info(msg):
    """弹窗提示，小白看得懂"""
    print(f"ℹ️ {msg}")
    try:
        Tk().messagebox.showinfo("提示", msg) #会弹出一个白色小窗口，标题写 “提示”，内容是你传进来的文字。
    except:
        pass

def show_error(msg):
    """错误弹窗"""
    print(f"❌ {msg}")
    try:
        Tk().messagebox.showerror("错误", msg)
    except:
        pass

# ====================== 主程序（极简封装） ======================
def main():
    print("=" * 60)
    print("🤖 HVE自动配置分配工具 - 小白一键版")
    print("📌 功能：自动读取BR → 生成HVE → 匹配CFG → 扣库存 → 填写S2F")
    print("=" * 60)
    print()

    # 1. 加载配置
    config = load_or_create_config()
    cfg = config["FILES"] #这里cfg是啥？是自动配置.ini里["FILES"]的整段内容的字典形式，即cfg = {
#     "s2f_file_path": "/xxx/xxx/S2F.xlsx",
#     "config_file_path": "/xxx/xxx/Config.xlsx",
#     "br_file_path": "/xxx/xxx/BR.xlsx"}
    sheet_cfg = config["SHEETS"]
    range_cfg = config["RANGE"]

    # 2. 自动检查文件，缺失则弹窗选择
    print("🔍 正在检查所需文件...")
    for key, name in [
        ("S2F_file_path", "S2F出货表"), #tuple "自动配置.ini"里对应变量名是小写那是parser自动的规则，不影响读取
        ("config_file_path", "Config配置表"),
        ("BR_file_path", "Build Rules规则表")
    ]:
        if not cfg[key] or not os.path.exists(cfg[key]): #前者是第一次，即"自动配置.ini"cfg[key]是空的时候；后者是文件不存在的时候，即非第一次，但是文件位置被动了
            show_info(f"请选择【{name}】Excel文件")
            fpath = select_excel_file(f"选择{name}")
            if not fpath:
                show_error(f"未选择{name}，程序退出")
                return #直接结束main程序
            cfg[key] = fpath

    # 保存配置
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        config.write(f) #在CONFIG_PATH地址的文件里再次写入更新后的内容（即把[FILES]里的值通过手动选取补充完整了）

    # 3. 读取参数
    try:
        BR_START = int(range_cfg["BR_start_row"])
        BR_END = int(range_cfg["BR_end_row"]) #所以如果要改不同Team的BR的首行和尾行，可以在自动配置文件里直接改
    except:
        BR_START, BR_END = 3, 14

    # 4. 打开所有Excel（自动处理）
    print("📂 正在打开所有Excel文件...")
    try:
        S2F_file = xl.load_workbook(cfg["S2F_file_path"])
        config_file = xl.load_workbook(cfg["config_file_path"])
        BR_file = xl.load_workbook(cfg["BR_file_path"])

        S2F_sheet = S2F_file[sheet_cfg["S2F_sheet_name"]]
        config_sheet = config_file[sheet_cfg["config_sheet_name"]]
        BR_HVE_sheet = BR_file[sheet_cfg["BR_HVE_sheet_name"]]
    except Exception as e: #Exception = 所有常见错误；as e = 把错误信息取名叫 e
        show_error(f"打开文件失败：{str(e)}")
        return

    # 5. 读取表头
    config_head = to_dic_head_col(config_sheet)
    S2F_head = to_dic_head_col(S2F_sheet)
    BR_head = to_dic_head_col(BR_HVE_sheet, head_row=2)

    # 6. 核心执行（全自动，无多余回车）
    total = BR_END - BR_START + 1
    print(f"\n🚀 开始自动处理（共 {total} 个HVE）")

    for idx, row in enumerate(range(BR_START, BR_END + 1), 1): #e.g., enumerate([5,6,7,8], 1) 打印：(1, 5)(2, 6)(3, 7)(4, 8)
        print(f"\n===== 【进度 {idx}/{total}】正在处理第 {row} 行 =====")

        # 生成HVE编码
        try:
            hve_code = cell_value_concatenator(
                cfg["BR_file_path"],
                sheet_cfg["BR_HVE_sheet_name"],
                cell1='D18',
                cell2=f"B{row}"
            )
            print(f"✅ 生成HVE编码：{hve_code}")
        except Exception as e:
            show_error(f"生成HVE失败：{e}")
            continue #直接结束本循环进入下一个，因为HVEcode都没有，后面都是扯淡

        # 生成订单数量
        try:
            hve_order = to_set_order(
                cfg["BR_file_path"],
                sheet_cfg["BR_HVE_sheet_name"],
                row, 15, hve_code
            )
            print(f"✅ 订单数量：{hve_order}")
        except Exception as e:
            show_error(f"读取订单失败：{e}")
            continue

        # 匹配CFG
        target_cfg = None
        target_cfgs = []
        for r in range(1, config_sheet.max_row + 1):
            alloc_val = config_sheet.cell(r, config_head['Allocations as']).value
            if alloc_val and to_normalize(hve_code) in to_lsplit_meaning(alloc_val):
                target_cfgs.append(config_sheet.cell(r, config_head['Name']).value)

        # 选择CFG
        if len(target_cfgs) > 1:
            print(f"⚠️ 找到多个CFG：{target_cfgs}")
            while True: #一直循环
                choose = input(f"👉 请输入要使用的CFG：").strip()
                if choose in target_cfgs:
                    target_cfg = choose
                    break
                print(f"❌ 输入错误，请从 {target_cfgs} 中选择")
        elif len(target_cfgs) == 1:
            target_cfg = target_cfgs[0]
            print(f"✅ 自动匹配CFG：{target_cfg}")
        else:
            print(f"❌ 未找到匹配的CFG，跳过本行")
            continue

        # 获取数量
        try:
            cfg_order = to_set_order(
                cfg["BR_file_path"],
                sheet_cfg["BR_HVE_sheet_name"],
                row, 15, target_cfg
            )
            need_qty = cfg_order[target_cfg]
            print(f"📦 需要分配数量：{need_qty}")
        except:
            print(f"❌ 获取数量失败，跳过")
            continue

        # 从Config表扣库存
        plate = out_taker(cfg["config_file_path"], sheet_cfg["config_sheet_name"], target_cfg, need_qty)
        if isinstance(plate, str):
            show_error(f"库存不足/找不到CFG：{plate}")
            continue
        print(f"✅ 成功取出 {need_qty} 个 {target_cfg}")

        # 填充到S2F
        filled = 0
        for s_row in range(1, S2F_sheet.max_row + 1):
            if plate[target_cfg] <= 0:
                break

            loc_val = S2F_sheet.cell(s_row, S2F_head['Location']).value
            br_cfg_val = BR_HVE_sheet.cell(row, BR_head['Config']).value
            s2f_cfg_val = S2F_sheet.cell(s_row, S2F_head['Config']).value

            if loc_val and br_cfg_val and not s2f_cfg_val:
                if to_normalize_cfg(loc_val) == to_normalize_cfg(br_cfg_val):
                    S2F_sheet.cell(s_row, S2F_head['Config']).value = target_cfg
                    plate[target_cfg] -= 1
                    filled += 1

        # 保存
        S2F_file.save(cfg["S2F_file_path"])
        print(f"✅ 已填充 {filled} 个位置 | S2F已自动保存")

    # 完成
    print("\n" + "=" * 60)
    show_info("🎉 全部HVE分配任务已完成！")
    print("✅ 程序运行结束")
    print("=" * 60)

# ====================== 一键启动 ======================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        show_error(f"程序崩溃：{str(e)}")
        with open("错误日志.txt", "a", encoding="utf-8") as f: #这里的a是啥模式？"a" = append，不会覆盖之前的错误
            f.write(f"错误：{str(e)}\n") #没有写任何路径，只写了文件名 → 默认就是当前文件夹
    input("\n按回车键退出...")