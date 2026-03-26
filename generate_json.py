# -*- coding: utf-8 -*-
import sys, os, codecs, json, glob
if os.name == 'nt':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

from openpyxl import load_workbook
from datetime import datetime

# 自动查找Downloads目录下最新的推广费Excel
def find_latest_excel():
    patterns = [
        os.path.join(r'C:\Users\15936\Downloads', '*推广费*.xlsx'),
        os.path.join(r'C:\Users\15936\Downloads', '*推广*.xlsx'),
    ]
    all_files = []
    for p in patterns:
        all_files.extend(glob.glob(p))
    if not all_files:
        return None
    return max(all_files, key=os.path.getmtime)

def pf(fee):
    if fee is None or fee == '/':
        return 0
    try:
        return float(str(fee).replace('%', ''))
    except:
        return 0

excel_path = find_latest_excel()
if not excel_path:
    print("ERROR: 未找到推广费Excel文件！")
    sys.exit(1)
print(f"使用Excel文件: {os.path.basename(excel_path)}")
wb = load_workbook(excel_path, data_only=True)
ws = wb['表1-费用明细']

# 读取表头，确定年份列
headers = [cell.value for cell in ws[2]]
# 总费用=col10, 第1年总计=col11, 第2年总计=col14, 第3年总计=col19, 第4年总计=col24, 第5年总计=col25
# 需要动态检测所有"第X年总计"列
year_cols = {}
for i, h in enumerate(headers):
    if h and '年总计' in str(h) and '第' in str(h):
        year_num = str(h).replace('第','').replace('年总计','')
        year_cols[year_num] = i + 1  # 1-based

years_list = sorted(year_cols.keys(), key=lambda x: int(x))
print(f"Detected year columns: {years_list}")

# 收集产品
products = {}

for row_idx in range(3, ws.max_row + 1):
    row = [cell.value for cell in ws[row_idx]]
    if not row[0]:
        continue
    
    prod_id = row[1]
    ins_type = str(row[2]).strip() if row[2] else ""
    prod_name = str(row[3]).strip() if row[3] else ""
    fee_rate = row[10]  # 总费用
    fee_combo = str(row[7]).strip() if row[7] else ""  # 费率组合
    attr = str(row[0]).strip()  # 商品属性
    
    key = f"{prod_id}_{prod_name}" if prod_id else prod_name
    if key not in products:
        # 找该产品所有变体中最大的总费用
        products[key] = {
            'id': prod_id,
            'name': prod_name,
            'type': ins_type,
            'attr': attr,
            'total_fee': fee_rate,
            'years': years_list,
            'variants': []
        }
    
    # 每个变体的各年总计
    yearly_totals = []
    for y in years_list:
        col_idx = year_cols[y]
        val = row[col_idx - 1] if col_idx <= len(row) else None  # 0-based
        if val is None or val == '/':
            yearly_totals.append('0')
        else:
            yearly_totals.append(str(val))
    
    # 总费用取所有变体中最高的
    current_max = pf(products[key]['total_fee'])
    if pf(fee_rate) > current_max:
        products[key]['total_fee'] = fee_rate
    
    products[key]['variants'].append({
        'fee_combo': fee_combo,
        'total_fee': str(fee_rate) if fee_rate else '0',
        'yearly_totals': yearly_totals,
    })

# 按总费用降序排序
sorted_products = sorted(products.values(), key=lambda x: pf(x['total_fee']), reverse=True)

output = {
    'updateTime': datetime.now().strftime('%Y-%m-%d %H:%M'),
    'totalCount': len(sorted_products),
    'products': sorted_products,
}

output_path = r'C:\Users\15936\WorkBuddy\20260326151224\niubao_products.json'
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"Done! {len(sorted_products)} products saved to {output_path}")
