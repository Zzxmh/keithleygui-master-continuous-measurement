# 油价对CPI与PPI的影响分析：直接效应与间接效应

## 研究目的
分别计算油价上升对中国、美国、日本和越南的CPI与PPI的**直接效应**和**间接效应**，使用2023年及之后的数据。

## 方法论与参考文献

### 直接效应（Direct Effects）
油价直接作用于能源相关价格：
- **CPI**：汽油、燃料油等能源消费品在消费篮子中的权重 × 油价弹性
- **PPI**：石油及石化产品作为中间投入的直接成本传递

**主要参考文献：**
1. Bachmeier, L., & Li, Q. (2007). *Pass-through of oil prices to domestic prices: Evidence from an oil-importing and an oil-exporting country.* Energy Economics.
2. Chen, S. S. (2009). *Oil price pass-through into inflation.* Energy Economics, 31(1), 126-133.
3. Cologni, A., & Manera, M. (2008). *Oil prices, inflation and interest rates in a structural cointegrated VAR model for the G-7 countries.* Energy Economics, 30(3), 856-888.

### 间接效应（Indirect Effects）
油价通过生产链条的传导：
- 运输成本上升 → 商品价格
- 石化原料成本 → 化工、塑料等中间品 → 最终产品
- 输入型通胀传导

**主要参考文献：**
1. Hamilton, J. D. (2003). *What is an oil shock?* Journal of Econometrics, 113(2), 363-398.
2. Nakajima, T. (2011). *Monetary policy transmission under zero interest rates: An extended time-varying parameter vector autoregression approach.* IMF Economic Review.
3. Hooker, M. A. (2002). *Are oil shocks inflationary? Asymmetric and nonlinear specifications versus changes in regime.* Journal of Money, Credit, and Banking.

### 估算方法
- **弹性法**：采用文献中的油价-通胀传递弹性系数
- **回归法**：基于2023+数据的月度VAR/OLS估计
- **权重法**：结合各国CPI/PPI篮子中能源权重

## 数据来源
- **油价**：WTI或Brent原油现货价格（美元/桶）
- **美国**：FRED (CPI-U, PPI)
- **中国**：国家统计局 / CEIC / Wind
- **日本**：日本总务省统计局 / OECD
- **越南**：越南统计局 / World Bank

## 使用说明
1. 安装依赖：`pip install -r requirements.txt`
2. （可选）在 `.env` 中设置 `FRED_API_KEY` 获取美国数据
3. 运行 `oil_cpi_ppi_analysis.ipynb`
