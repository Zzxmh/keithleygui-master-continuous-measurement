"""
CPI/PPI与油价数据加载模块
支持FRED API及CSV文件加载，筛选2023年及之后数据
"""

import os
from datetime import datetime
import pandas as pd
import numpy as np

# FRED数据系列代码 (详见 https://fred.stlouisfed.org/)
FRED_SERIES = {
    "oil_wti": "DCOILWTICO",      # WTI原油现货价格
    "oil_brent": "DCOILBRENTEU",  # Brent原油
    "us_cpi": "CPIAUCSL",         # 美国CPI-U
    "us_ppi": "PPIACO",           # 美国PPI (全部商品)
}


def load_fred_data(api_key: str = None, start_date: str = "2023-01-01") -> pd.DataFrame:
    """
    从FRED加载美国及油价数据
    
    Parameters:
    -----------
    api_key : str, optional
        FRED API密钥，可从 https://fred.stlouisfed.org/docs/api/api_key.html 申请
    start_date : str
        起始日期，默认2023-01-01
    """
    try:
        from fredapi import Fred
        key = api_key or os.environ.get("FRED_API_KEY")
        if not key:
            raise ValueError("需要FRED_API_KEY，请设置环境变量或在调用时传入")
        fred = Fred(api_key=key)
        
        dfs = []
        for name, series_id in FRED_SERIES.items():
            try:
                s = fred.get_series(series_id, observation_start=start_date)
                s.name = name
                dfs.append(s)
            except Exception as e:
                print(f"警告: 无法获取 {name} ({series_id}): {e}")
        
        if not dfs:
            return pd.DataFrame()
        df = pd.concat(dfs, axis=1)
        df = df.dropna(how='all')
        return df
    except ImportError:
        print("未安装fredapi，请运行: pip install fredapi")
        return pd.DataFrame()


def load_sample_data() -> pd.DataFrame:
    """
    加载示例/模拟数据（当无API时使用）
    基于2023-2024年典型月度数据构造
    """
    # 2023-2024 典型油价 (WTI, 美元/桶) 月度平均
    oil_data = {
        "2023-01": 81.0, "2023-02": 76.0, "2023-03": 73.0, "2023-04": 81.0,
        "2023-05": 72.0, "2023-06": 70.0, "2023-07": 76.0, "2023-08": 82.0,
        "2023-09": 90.0, "2023-10": 86.0, "2023-11": 78.0, "2023-12": 72.0,
        "2024-01": 72.0, "2024-02": 77.0, "2024-03": 81.0, "2024-04": 85.0,
        "2024-05": 77.0, "2024-06": 78.0, "2024-07": 82.0, "2024-08": 81.0,
        "2024-09": 83.0, "2024-10": 79.0, "2024-11": 76.0, "2024-12": 74.0,
    }
    
    # 美国CPI (2023-01=300为基准)
    us_cpi = {
        "2023-01": 300.0, "2023-02": 300.8, "2023-03": 302.1, "2023-04": 303.4,
        "2023-05": 304.1, "2023-06": 304.2, "2023-07": 305.4, "2023-08": 307.0,
        "2023-09": 307.8, "2023-10": 307.7, "2023-11": 307.5, "2023-12": 306.7,
        "2024-01": 308.4, "2024-02": 310.2, "2024-03": 312.2, "2024-04": 313.5,
        "2024-05": 314.1, "2024-06": 314.1, "2024-07": 314.6, "2024-08": 314.3,
        "2024-09": 315.0, "2024-10": 315.5, "2024-11": 315.8, "2024-12": 316.0,
    }
    
    # 美国PPI (2023-01=190)
    us_ppi = {
        "2023-01": 190.0, "2023-02": 189.5, "2023-03": 188.2, "2023-04": 187.5,
        "2023-05": 186.8, "2023-06": 186.0, "2023-07": 187.2, "2023-08": 188.5,
        "2023-09": 189.8, "2023-10": 189.0, "2023-11": 188.0, "2023-12": 187.2,
        "2024-01": 188.5, "2024-02": 189.8, "2024-03": 191.0, "2024-04": 191.5,
        "2024-05": 191.0, "2024-06": 190.5, "2024-07": 191.2, "2024-08": 191.0,
        "2024-09": 191.5, "2024-10": 191.2, "2024-11": 190.8, "2024-12": 190.5,
    }
    
    dates = pd.date_range(start="2023-01-01", periods=24, freq="MS")
    df = pd.DataFrame({
        "oil_wti": [oil_data.get(d.strftime("%Y-%m"), np.nan) for d in dates],
        "us_cpi": [us_cpi.get(d.strftime("%Y-%m"), np.nan) for d in dates],
        "us_ppi": [us_ppi.get(d.strftime("%Y-%m"), np.nan) for d in dates],
    }, index=dates)
    df.index.name = "date"
    return df


def load_country_data_from_csv(data_dir: str = "data") -> dict:
    """
    从CSV加载各国CPI/PPI数据
    期望文件: {data_dir}/china_cpi.csv, usa_cpi.csv, japan_cpi.csv, vietnam_cpi.csv
    以及对应的 _ppi.csv，以及 oil_prices.csv
    """
    result = {}
    if not os.path.exists(data_dir):
        return result
    
    for country in ["china", "usa", "japan", "vietnam"]:
        for var in ["cpi", "ppi"]:
            fpath = os.path.join(data_dir, f"{country}_{var}.csv")
            if os.path.exists(fpath):
                df = pd.read_csv(fpath, parse_dates=True, index_col=0)
                result[f"{country}_{var}"] = df
    
    oil_path = os.path.join(data_dir, "oil_prices.csv")
    if os.path.exists(oil_path):
        result["oil"] = pd.read_csv(oil_path, parse_dates=True, index_col=0)
    
    return result
