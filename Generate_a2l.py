import re
from pathlib import Path
from header_template import A2L_HEADER
from Characteristic_rule import A2L_CHARA_RULE

def hex_addr_to_a2l(addr_hex):
    # 十六进制字符串转十进制，乘2，再转回十六进制字符串
    addr_int = int(addr_hex, 16)
    addr_int2 = addr_int * 2
    return f'0x{addr_int2:X}'

# 处理name(如果开头是_，则去掉)
def process_name(name):
    if name.startswith('_'):
        return name[1:]
    return name

def process_map_to_a2l(map_path: Path, out_path: Path) -> None:
    # 如果已存在同名文件，直接覆写（'wt'模式等价于'w'，始终覆写）
    inside_table = False
    with open(map_path, encoding='utf-8', errors='ignore') as f_map, \
         open(out_path, 'wt', encoding='utf-8') as f_a2l:
        # 添加头部信息
        f_a2l.write(A2L_HEADER)

        for line in f_map:
            # 定位起始标记
            if "GLOBAL SYMBOLS: SORTED BY Symbol Address" in line:
                inside_table = True
                next(f_map)          # 跳过表头虚线行
                continue

            if not inside_table:
                continue

            # 去掉首尾空白后按空格/制表符切分
            parts = line.strip().split()
            # 结束条件：切不出 3 列，或地址/名称为空
            if len(parts) < 3:
                break
            page = parts[0]
            addr_hex = parts[1]
            name = parts[2]

            # 判断是否为特殊CHARACTERISTIC格式
            # 定义匹配前缀
            chara_prefixes = ['_CFG_', '_COUT_', '_DINP_']
            # is_chara = False

            for prefix in chara_prefixes:
                if name.startswith(prefix):
                    parts_ = name.split('_', 2)
                    if len(parts_) == 3:
                        after_ = parts_[2]
                        ecu_addr = hex_addr_to_a2l(addr_hex)

                    # 后缀为b
                    if after_.startswith('b'):
                        block = (
                            f"    /begin CHARACTERISTIC {process_name(name)} \"\"\n"
                            f"      VALUE {ecu_addr} Record_BOOLEAN 0 NO_COMPU_METHOD 0 1\n"
                            f"    /end CHARACTERISTIC\n\n"
                        )
                        f_a2l.write(block)
                        # is_chara = True
                        break
                    
                    # 后缀为st
                    elif after_.startswith('st'):
                        block = (
                            f"    /begin CHARACTERISTIC {process_name(name)} \"\"\n"
                            f"      VALUE {ecu_addr} Record_UWORD 0 NO_COMPU_METHOD 0 65535\n"
                            f"    /end CHARACTERISTIC\n\n"
                        )
                        f_a2l.write(block)
                        # is_chara = True
                        break

            # ---------- 如果没命中 CHARACTERISTIC，再看是否生成 MEASUREMENT ----------
            else:   # 注意：这里的 else 对应的是 for 循环，不是 if
                # name筛选：只能以单个下划线开头，且后面只能是字母数字下划线
                # 检查地址是否为合法十六进制
                if re.fullmatch(r'_[A-Za-z][A-Za-z0-9_]*', name) and re.fullmatch(r'[0-9A-Fa-f]+', addr_hex):
                    ecu_addr = hex_addr_to_a2l(addr_hex)
                    block = (
                        f"    /begin MEASUREMENT {process_name(name)} \"\"\n"
                        f"      UWORD NO_COMPU_METHOD 0 0 0 65535\n"
                        f"      ECU_ADDRESS {ecu_addr}\n"
                        f"    /end MEASUREMENT\n\n"
                    )
                    f_a2l.write(block)
            # 写完所有 MEASUREMENT，补结尾
        f_a2l.write(A2L_CHARA_RULE)
        f_a2l.write("\n\n")
        f_a2l.write("  /end MODULE\n")
        f_a2l.write("/end PROJECT\n")

# 路径替换为你的实际文件路径
map_path = r'c:\Users\xiehaoyang\Desktop\生成A2L文件\evcu2022_20250103_0919.map'
out_path = r'c:\Users\xiehaoyang\Desktop\生成A2L文件\output_measurement.a2l'

process_map_to_a2l(map_path, out_path)
print('已生成output_measurement.a2l文件')










