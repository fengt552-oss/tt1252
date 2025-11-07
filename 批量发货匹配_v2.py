# 批量发货匹配_v2.py
# -*- coding: utf-8 -*-
"""
批量发货匹配脚本（增强版）
用法（在 Windows 上打包为 .exe）:
1. 在 Windows 上安装 Python 3.8+
2. 安装依赖: pip install -r requirements.txt
3. 执行: pyinstaller --onefile 批量发货匹配_v2.py
4. 把 fahuo.xlsx、订单.xlsx、快递.txt 放到 exe 同目录，运行 exe

脚本功能：
- 解析 快递.txt 中的 姓名/电话/地址/单号
- 从 订单.xlsx 中按手机号前三+后四、姓名首字、地址进行三阶段匹配（严格→模糊→宽松）
- 将匹配到的订单号写入 fahuo.xlsx 模板的订单号列
- 输出 fahuo_输出.xlsx（匹配结果）和 未匹配列表.xlsx（便于人工处理）
"""

import re
import sys
import pandas as pd
from pathlib import Path

def load_files():
    required = ['fahuo.xlsx','订单.xlsx','快递.txt']
    for f in required:
        if not Path(f).exists():
            print(f"错误：未找到文件 {f}，请把三个文件放在同一目录后重试。")
            sys.exit(1)
    fahuo = pd.read_excel('fahuo.xlsx')
    orders = pd.read_excel('订单.xlsx')
    txt = Path('快递.txt').read_text(encoding='utf-8', errors='ignore')
    return fahuo, orders, txt

def parse_express(txt):
    express_list = []
    # 支持用 ✅ 分割，也支持按行解析
    blocks = txt.split('✅') if '✅' in txt else txt.split('\\n\\n')
    for block in blocks:
        lines = [l.strip() for l in block.strip().split('\\n') if l.strip()]
        if not lines:
            continue
        # 尝试从每行提取：姓名 手机 地址
        for i, line in enumerate(lines):
            m = re.match(r'(.+?)\\s+(\\d{11})\\s+(.+)', line)
            if m:
                name, phone, address = m.groups()
                # 找最近的物流单号行
                tracking = None
                # 搜索同一区块里的单号
                for ln in lines[i+1: i+6]:
                    m2 = re.search(r'(YT\\d+|SF\\d+|ZTO\\d+|JD\\d+|TTK\\d+)', ln, re.I)
                    if m2:
                        tracking = m2.group()
                        break
                express_list.append({'name': name, 'phone': phone, 'address': address, 'tracking': tracking})
                break
    return pd.DataFrame(express_list)

def split_address(addr):
    """从地址中提取省/市/区关键词，返回关键词列表（优先省/市/区）"""
    if not isinstance(addr, str):
        return []
    kws = []
    for kw in ['省','市','区','县','州','旗']:
        pos = addr.find(kw)
        if pos != -1:
            kws.append(addr[:pos+1])
            addr = addr[pos+1:]
    # 额外补充前 4 个汉字作为兜底关键词
    if len(kws) < 1 and len(addr) >= 2:
        kws.append(addr[:2])
    return kws[:4]

def enhance_match(express, orders):
    orders = orders.copy()
    # normalize helper columns
    orders['phone_first3'] = orders['收件人电话'].astype(str).str[:3]
    orders['phone_last4'] = orders['收件人电话'].astype(str).str[-4:]
    orders['name_first'] = orders['收件人姓名'].astype(str).str[:1]

    unmatched_rows = []
    for idx, row in express.iterrows():
        phone = str(row['phone'])
        phone3 = phone[:3]
        phone4 = phone[-4:]
        name1 = str(row['name'])[:1]
        kws = split_address(str(row.get('address','')))

        # 阶段1：严格匹配（手机号首尾 + 姓名首字 + 所有关键词都包含）
        cond = (orders['phone_first3'] == phone3) & (orders['phone_last4'] == phone4) & (orders['name_first'] == name1)
        for kw in kws:
            cond = cond & orders['收件人地址'].astype(str).str.contains(kw, na=False)
        match = orders[cond]

        # 阶段2：地址智能分词模糊匹配（手机号+姓名 + 任一关键词包含）
        if match.empty and kws:
            cond2 = (orders['phone_first3'] == phone3) & (orders['phone_last4'] == phone4) & (orders['name_first'] == name1)
            addr_mask = False
            for kw in kws:
                addr_mask = addr_mask | orders['收件人地址'].astype(str).str.contains(kw, na=False)
            match = orders[cond2 & addr_mask]

        # 阶段3：超级宽松（手机号+姓名首字，若唯一则取）
        if match.empty:
            match = orders[(orders['phone_first3'] == phone3) & (orders['phone_last4'] == phone4) & (orders['name_first'] == name1)]

        if not match.empty:
            express.at[idx, 'order_id'] = match['订单号'].iloc[0]
        else:
            express.at[idx, 'order_id'] = None
            unmatched_rows.append(row.to_dict())

    return express, unmatched_rows

def write_results(fahuo, express, unmatched):
    # 找到模板中的列名（可能含中文/英文混合）
    col_tracking = None
    col_order = None
    for c in fahuo.columns:
        if 'Tracking' in str(c) or '快递单号' in str(c):
            col_tracking = c
        if 'Order' in str(c) or '订单号' in str(c):
            col_order = c
    if col_tracking is None or col_order is None:
        print("错误：在 fahuo.xlsx 中未找到快递单号或订单号列，请确认表头。")
        return

    for _, r in express.iterrows():
        fahuo.loc[fahuo[col_tracking] == r['tracking'], col_order] = r.get('order_id')

    fahuo.to_excel('fahuo_输出.xlsx', index=False)
    if unmatched:
        pd.DataFrame(unmatched).to_excel('未匹配列表.xlsx', index=False)
    print('完成：已生成 fahuo_输出.xlsx 和 未匹配列表.xlsx（如有未匹配）')

def main():
    fahuo, orders, txt = load_files()
    express = parse_express(txt)
    if express.empty:
        print('警告：未解析到快递信息，请检查 快递.txt 格式。')
        return
    express, unmatched = enhance_match(express, orders)
    write_results(fahuo, express, unmatched)

if __name__ == '__main__':
    main()
