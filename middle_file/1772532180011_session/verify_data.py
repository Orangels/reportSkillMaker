#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最终数据准确性验证
"""

# 验证数据关系
effective = 3013
harass = 1
total = 3014

criminal = 18
security = 157
traffic = 923
dispute = 281
emergency = 858
other = 777

sum_categories = criminal + security + traffic + dispute + emergency + other
print(f"各大类之和: {sum_categories}, 有效警情: {effective}, 匹配: {sum_categories == effective}")
print(f"有效警情+骚扰: {effective + harass}, 总接报: {total}, 匹配: {effective + harass == total}")

# 环比计算验证
import json

def verify_pct(name, this_month, last_month, claimed_pct):
    actual_pct = round((this_month - last_month) / last_month * 100, 2)
    match = abs(actual_pct - claimed_pct) < 0.1
    status = "PASS" if match else "FAIL"
    print(f"  [{status}] {name}: 本月{this_month}, 上月{last_month}, 计算环比{actual_pct}%, 数据环比{claimed_pct}%")

print("\n环比变化率验证:")
verify_pct("有效警情", 3013, 2702, 11.51)
verify_pct("刑事警情", 18, 24, -25.0)
verify_pct("治安警情", 157, 164, -4.27)
verify_pct("交通警情", 923, 800, 15.38)
verify_pct("纠纷警情", 281, 260, 8.08)
verify_pct("群众紧急求助", 858, 793, 8.2)
verify_pct("其他警情", 777, 663, 17.19)

# 验证上升/下降分类正确性
print("\n上升/下降分类验证:")
rising = {"交通警情": 15.38, "纠纷警情": 8.08, "群众紧急求助": 8.2, "其他警情": 17.19}
falling = {"刑事警情": -25.0, "治安警情": -4.27}

for name, pct in rising.items():
    status = "PASS" if pct > 0 else "FAIL"
    print(f"  [{status}] {name} 环比{pct}% (归入上升)")

for name, pct in falling.items():
    status = "PASS" if pct < 0 else "FAIL"
    print(f"  [{status}] {name} 环比{pct}% (归入下降)")

print("\n所有验证完成。")
