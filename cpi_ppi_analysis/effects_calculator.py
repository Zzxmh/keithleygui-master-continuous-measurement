"""
油价对CPI/PPI的直接效应与间接效应估算模块

基于文献弹性系数及回归方法，分别计算直接效应和间接效应。
参考文献见 README.md
"""

import pandas as pd
import numpy as np
from typing import Tuple, Dict, Optional


# 各国CPI篮子中能源权重 (近似值，基于OECD/各国统计局)
# 来源: OECD CPI weights, 各国统计年鉴
ENERGY_WEIGHTS_CPI = {
    "USA": 0.072,    # 美国能源约7.2%
    "China": 0.025,  # 中国能源权重较低
    "Japan": 0.078,  # 日本能源依赖进口
    "Vietnam": 0.035,
}

# 各国PPI中石油/能源相关权重 (近似)
ENERGY_WEIGHTS_PPI = {
    "USA": 0.12,
    "China": 0.08,
    "Japan": 0.10,
    "Vietnam": 0.06,
}

# 文献中的油价-价格传递弹性 (直接效应)
# Bachmeier & Li (2007), Chen (2009) 等
# 油价上涨1% -> CPI/PPI变动百分比
DIRECT_ELASTICITY_CPI = {
    "USA": 0.15,     # 美国CPI对油价弹性
    "China": 0.08,   # 中国管制较多，弹性较低
    "Japan": 0.18,   # 日本能源依赖进口
    "Vietnam": 0.12,
}

DIRECT_ELASTICITY_PPI = {
    "USA": 0.25,
    "China": 0.20,
    "Japan": 0.28,
    "Vietnam": 0.22,
}

# 间接效应弹性 (通过投入产出链)
# Hamilton (2003), Hooker (2002) 等
INDIRECT_ELASTICITY_CPI = {
    "USA": 0.06,
    "China": 0.05,
    "Japan": 0.07,
    "Vietnam": 0.08,  # 新兴市场间接传导可能更强
}

INDIRECT_ELASTICITY_PPI = {
    "USA": 0.12,
    "China": 0.15,    # 中国制造业链长
    "Japan": 0.10,
    "Vietnam": 0.14,
}


def compute_direct_effect(
    oil_pct_change: float,
    country: str,
    var: str = "CPI"
) -> float:
    """
    计算直接效应：油价变动对CPI/PPI的直接影响
    
    Parameters:
    -----------
    oil_pct_change : float
        油价变动百分比 (例如 10 表示上涨10%)
    country : str
        国家代码: USA, China, Japan, Vietnam
    var : str
        'CPI' 或 'PPI'
    
    Returns:
    --------
    float : 直接效应导致的CPI/PPI变动百分比
    """
    country = country.strip().title()
    if country == "Usa":
        country = "USA"
    if var.upper() == "CPI":
        elast = DIRECT_ELASTICITY_CPI.get(country, 0.12)
        weight = ENERGY_WEIGHTS_CPI.get(country, 0.05)
    else:
        elast = DIRECT_ELASTICITY_PPI.get(country, 0.22)
        weight = ENERGY_WEIGHTS_PPI.get(country, 0.09)
    
    # 直接效应 = 权重 * 弹性 * 油价变动
    direct = weight * elast * (oil_pct_change / 100.0) * 100
    return direct


def compute_indirect_effect(
    oil_pct_change: float,
    country: str,
    var: str = "CPI"
) -> float:
    """
    计算间接效应：油价通过生产链对CPI/PPI的间接影响
    """
    country = country.strip().title()
    if country == "Usa":
        country = "USA"
    if var.upper() == "CPI":
        elast = INDIRECT_ELASTICITY_CPI.get(country, 0.06)
    else:
        elast = INDIRECT_ELASTICITY_PPI.get(country, 0.12)
    
    indirect = elast * (oil_pct_change / 100.0) * 100
    return indirect


def estimate_elasticity_from_data(
    oil_series: pd.Series,
    price_series: pd.Series,
    lag: int = 1
) -> Tuple[float, float]:
    """
    从数据估计油价-价格弹性 (用于验证或替代文献系数)
    
    Returns:
    --------
    (direct_elast, indirect_elast) : 直接与间接弹性估计
    """
    try:
        from statsmodels.regression.linear_model import OLS
        from statsmodels.tools import add_constant
        
        df = pd.concat([oil_series, price_series], axis=1).dropna()
        if len(df) < 12:
            return np.nan, np.nan
        
        # 对数差分
        oil_log = np.log(df.iloc[:, 0])
        price_log = np.log(df.iloc[:, 1])
        
        doil = oil_log.diff().dropna()
        dprice = price_log.diff().dropna()
        common = doil.index.intersection(dprice.index)
        doil, dprice = doil.loc[common], dprice.loc[common]
        
        # 加入滞后期捕捉间接效应
        if lag >= 1 and len(doil) > lag:
            doil_lag = doil.shift(lag).dropna()
            common2 = doil_lag.index.intersection(dprice.index)
            X = pd.concat([doil.loc[common2], doil_lag.loc[common2]], axis=1)
            X = add_constant(X.fillna(0))
            y = dprice.loc[common2]
            model = OLS(y, X).fit()
            direct_elast = model.params.iloc[1] if len(model.params) > 1 else np.nan
            indirect_elast = model.params.iloc[2] if len(model.params) > 2 else 0
        else:
            X = add_constant(doil)
            y = dprice.loc[doil.index]
            model = OLS(y, X).fit()
            direct_elast = model.params.iloc[1]
            indirect_elast = 0
        
        return float(direct_elast), float(indirect_elast)
    except Exception:
        return np.nan, np.nan


def compute_all_effects(
    oil_pct_change: float,
    countries: list = None
) -> pd.DataFrame:
    """
    计算所有国家、CPI/PPI的直接与间接效应
    
    Parameters:
    -----------
    oil_pct_change : float
        油价上涨百分比 (如 10 = 10%)
    countries : list
        国家列表，默认 ["USA", "China", "Japan", "Vietnam"]
    
    Returns:
    --------
    pd.DataFrame : 结果表格
    """
    countries = countries or ["USA", "China", "Japan", "Vietnam"]
    rows = []
    for country in countries:
        for var in ["CPI", "PPI"]:
            direct = compute_direct_effect(oil_pct_change, country, var)
            indirect = compute_indirect_effect(oil_pct_change, country, var)
            total = direct + indirect
            rows.append({
                "国家": country,
                "指标": var,
                "直接效应(%)": round(direct, 4),
                "间接效应(%)": round(indirect, 4),
                "总效应(%)": round(total, 4),
                "油价变动(%)": oil_pct_change,
            })
    return pd.DataFrame(rows)


def compute_effects_from_series(
    oil_series: pd.Series,
    price_dict: Dict[str, Dict[str, pd.Series]],
    base_date: Optional[str] = None
) -> pd.DataFrame:
    """
    基于实际数据序列计算效应
    
    price_dict 格式: {"USA": {"CPI": series, "PPI": series}, ...}
    """
    if base_date:
        oil_series = oil_series[oil_series.index >= base_date]
    
    oil_pct = oil_series.pct_change().dropna() * 100
    if oil_pct.empty:
        oil_avg_change = 0
    else:
        oil_avg_change = float(oil_pct.mean())
    
    results = []
    for country, vars_dict in price_dict.items():
        for var, ps in vars_dict.items():
            direct = compute_direct_effect(oil_avg_change, country, var)
            indirect = compute_indirect_effect(oil_avg_change, country, var)
            results.append({
                "国家": country,
                "指标": var,
                "直接效应(%)": round(direct, 4),
                "间接效应(%)": round(indirect, 4),
                "总效应(%)": round(direct + indirect, 4),
                "样本期油价平均变动(%)": round(oil_avg_change, 2),
            })
    return pd.DataFrame(results)
