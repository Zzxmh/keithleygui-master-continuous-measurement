#!/usr/bin/env python3
"""
批量运行油价-CPI/PPI效应分析并输出结果
使用2023年及之后的数据，基于文献弹性系数
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from effects_calculator import compute_all_effects, compute_effects_from_series
from data_loader import load_fred_data, load_sample_data
import pandas as pd

START_DATE = "2023-01-01"


def main():
    print("=" * 70)
    print("油价对CPI与PPI的影响：直接效应与间接效应（中国、美国、日本、越南）")
    print("数据期：2023年及之后 | 方法论：文献弹性系数")
    print("=" * 70)

    # 1. 加载数据
    df = load_fred_data(start_date=START_DATE)
    if df.empty or len(df) < 6:
        data_path = os.path.join(os.path.dirname(__file__), "data", "sample_data_2023.csv")
        df = pd.read_csv(data_path, parse_dates=["date"], index_col="date")
        print("\n[数据] 使用样本数据 (2023-2024)")
    else:
        print("\n[数据] 已从FRED加载")

    df = df[df.index >= START_DATE]
    print(f"数据期: {df.index.min().strftime('%Y-%m')} 至 {df.index.max().strftime('%Y-%m')}\n")

    # 2. 情景分析
    for pct in [10, 20]:
        print(f"\n--- 情景：油价上涨 {pct}% ---")
        r = compute_all_effects(pct)
        print(r.to_string(index=False))

    # 3. 基于实际数据
    oil_col = "oil_wti" if "oil_wti" in df.columns else df.columns[0]
    price_dict = {}
    for country, ckey, pkey in [
        ("USA", "us_cpi", "us_ppi"),
        ("USA", "usa_cpi", "usa_ppi"),
        ("China", "china_cpi", "china_ppi"),
        ("Japan", "japan_cpi", "japan_ppi"),
        ("Vietnam", "vietnam_cpi", "vietnam_ppi"),
    ]:
        if ckey in df.columns or pkey in df.columns:
            price_dict[country] = {}
            if ckey in df.columns:
                price_dict[country]["CPI"] = df[ckey]
            if pkey in df.columns:
                price_dict[country]["PPI"] = df[pkey]
            if country not in price_dict or not price_dict[country]:
                price_dict.pop(country, None)

    # 去重USA
    if "USA" in price_dict and ("us_cpi" in df.columns or "usa_cpi" in df.columns):
        price_dict["USA"] = {}
        if "us_cpi" in df.columns:
            price_dict["USA"]["CPI"] = df["us_cpi"]
        elif "usa_cpi" in df.columns:
            price_dict["USA"]["CPI"] = df["usa_cpi"]
        if "us_ppi" in df.columns:
            price_dict["USA"]["PPI"] = df["us_ppi"]
        elif "usa_ppi" in df.columns:
            price_dict["USA"]["PPI"] = df["usa_ppi"]

    if price_dict:
        print("\n--- 基于样本期实际油价平均变动的效应 ---")
        r_act = compute_effects_from_series(df[oil_col], price_dict, base_date=START_DATE)
        print(r_act.to_string(index=False))

    print("\n" + "=" * 70)
    print("参考文献: Bachmeier & Li (2007), Chen (2009), Hamilton (2003), Cologni & Manera (2008)")
    print("=" * 70)


if __name__ == "__main__":
    main()
