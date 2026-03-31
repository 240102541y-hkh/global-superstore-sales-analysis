import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from prophet import Prophet   # 新增：需要先运行 pip install prophet
from datetime import datetime

# ==================== 配置 ====================
plt.rcParams["font.sans-serif"] = ["SimHei"]  # 设置字体为黑体（或 "Microsoft YaHei"）
plt.rcParams["axes.unicode_minus"] = False    # 解决负号显示成方框的问题

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==================== 1. 数据加载与清洗 ====================
file = "Global Superstore Data.xlsx"

orders = pd.read_excel(file, sheet_name="Orders")
returns = pd.read_excel(file, sheet_name="Returns")
people = pd.read_excel(file, sheet_name="People")

# 去重
returns = returns.drop_duplicates().reset_index(drop=True)

# 日期转换
orders["Order Date"] = pd.to_datetime(orders["Order Date"])
orders["Ship Date"] = pd.to_datetime(orders["Ship Date"])

# 新增发货耗时
orders["Shipping_Days"] = (orders["Ship Date"] - orders["Order Date"]).dt.days

# 合并 Returns
returns_selected = returns[["Order ID", "Returned"]]
data = orders.merge(returns_selected, on="Order ID", how="left")
data["Returned"] = data["Returned"].fillna(0)
data["Returned"] = data["Returned"].map(lambda x: 1 if x == "Yes" else 0).astype(int)

# 合并 People（重要！简历亮点）
data = data.merge(people, on="Region", how="left")

# ==================== 2. 基础信息 ====================
print("=== 数据基本信息 ===")
print(f"总订单数: {len(data):,}")
print(f"总销售额: ${data['Sales'].sum():,.2f}")
print(f"总利润: ${data['Profit'].sum():,.2f}")
print(f"整体退货率: {data['Returned'].mean():.2%}\n")

# ==================== 3. 时间序列分析 ====================
data["Year"] = data["Order Date"].dt.year
data["Month"] = data["Order Date"].dt.month
data["Year_Month"] = data["Year"].astype(str) + "-" + data["Month"].astype(str).str.zfill(2)

monthly_sales = data.groupby(["Year", "Month", "Year_Month"])["Sales"].sum().reset_index()
monthly_sales = monthly_sales.sort_values(["Year", "Month"])

# 保存月度销售趋势图
plt.figure(figsize=(14, 6))
sns.lineplot(data=monthly_sales, x="Year_Month", y="Sales", marker="o", linewidth=2.5)
plt.title("全球超市月度销售额趋势 (2014-2017)", fontsize=16, pad=20)
plt.xlabel("时间")
plt.ylabel("销售额 ($)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/1_monthly_sales_trend.png", dpi=300, bbox_inches='tight')
plt.close()
print("✅ 月度销售趋势图已保存")

# ==================== 4. 区域销售分析 ====================
region_sales = data.groupby("Region")["Sales"].sum().sort_values(ascending=False).reset_index()

plt.figure(figsize=(12, 6))
sns.barplot(y="Region", x="Sales", hue="Region", data=region_sales, palette="viridis", legend=False)
#sns.barplot(data=region_sales, x="Sales", y="Region", palette="Blues_d")
plt.title("各地区总销售额", fontsize=16)
plt.xlabel("销售额 ($)")
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/2_region_sales.png", dpi=300, bbox_inches='tight')
plt.close()

# ==================== 5. 品类利润分析 ====================
category_profit = data.groupby("Category")["Profit"].sum().sort_values(ascending=False).reset_index()
subcat_profit = data.groupby("Sub-Category")["Profit"].sum().sort_values(ascending=False).reset_index()

# 最赚钱 & 最亏钱的子类别
print("=== Top 5 最赚钱子类别 ===")
print(subcat_profit.head())
print("\n=== Top 5 最亏损子类别 ===")
print(subcat_profit.tail())

# ==================== 6. 退货率分析 ====================
return_rate = data.groupby("Category")["Returned"].mean().reset_index()
return_rate["Return Rate (%)"] = return_rate["Returned"] * 100

plt.figure(figsize=(10, 5))
sns.barplot(
    data=return_rate,
    x="Category",
    y="Return Rate (%)",
    hue="Category",  # 【新增】和 x 变量保持一致
    palette="Reds_d",
    legend=False     # 【新增】隐藏多余的图例
)
plt.title("各品类退货率", fontsize=16)
plt.ylabel("退货率 (%)")
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/3_return_rate_by_category.png", dpi=300, bbox_inches='tight')
plt.close()

# ==================== 7. 额外高阶分析（简历加分项）===================
# 利润率
data["Profit_Margin"] = data["Profit"] / data["Sales"] * 100
print(f"平均利润率: {data['Profit_Margin'].mean():.2f}%")

# 发货时长对退货的影响
shipping_return = data.groupby("Shipping_Days")["Returned"].mean().reset_index()
print("\n发货天数与退货率关系（示例）:")
print(shipping_return.head(10))

print(f"\n🎉 所有分析图表已保存到 {OUTPUT_DIR} 文件夹！")
print("建议：把这些图直接放进简历 PDF，或者做成 Streamlit/ PowerBI 仪表盘，会非常加分！")

print("=== 基础分析已完成 ===")

# ==================== 8. RFM 客户分层（高价值客户分析）===================
print("\n=== 8. RFM 客户分层分析 ===")
latest_date = data["Order Date"].max()

rfm = data.groupby("Customer ID").agg({
    "Order Date": lambda x: (latest_date - x.max()).days,      # Recency（距今多少天）
    "Order ID": "nunique",                                     # Frequency（下单次数）
    "Sales": "sum"                                             # Monetary（总消费金额）
}).reset_index()

rfm.columns = ["Customer ID", "Recency", "Frequency", "Monetary"]

# ----------------------
# 【关键修复】RFM 打分（解决分箱边界重复问题）
# ----------------------
r_labels = range(5, 0, -1)
f_labels = range(1, 6)
m_labels = range(1, 6)

# 1. Recency 打分（通常重复值少，直接用 qcut）
rfm["R_Score"] = pd.qcut(rfm["Recency"], q=5, labels=r_labels, duplicates='drop').astype(int)

# 2. Frequency 打分（核心修复：两种方案选其一）
# === 方案一：快速修复（加 duplicates='drop'，自动删除重复边界）===
# rfm["F_Score"] = pd.qcut(rfm["Frequency"], q=5, labels=f_labels, duplicates='drop').astype(int)

# === 方案二：手动分箱（更灵活，推荐用于论文，需先看数据分布）===
# 先看一下 Frequency 的分布：print(rfm["Frequency"].describe())
# 假设分布是：大部分人下单 1-2 次，少数人下单 3+ 次
f_bins = [0, 1, 2, 4, 7, float('inf')]  # 手动定义边界
rfm["F_Score"] = pd.cut(rfm["Frequency"], bins=f_bins, labels=f_labels).astype(int)

# 3. Monetary 打分（通常重复值少，直接用 qcut）
rfm["M_Score"] = pd.qcut(rfm["Monetary"], q=5, labels=m_labels, duplicates='drop').astype(int)

# ----------------------
# 后续代码（保持不变，顺便修复了 sns 的警告）
# ----------------------
rfm["RFM_Score"] = rfm["R_Score"].astype(str) + rfm["F_Score"].astype(str) + rfm["M_Score"].astype(str)
rfm["RFM_Total"] = rfm["R_Score"] + rfm["F_Score"] + rfm["M_Score"]

# 分层逻辑
def rfm_segment(row):
    if row["RFM_Total"] >= 12: return "高价值客户"
    elif row["RFM_Total"] >= 9: return "中价值客户"
    elif row["RFM_Total"] >= 6: return "一般客户"
    else: return "流失风险客户"

rfm["Segment"] = rfm.apply(rfm_segment, axis=1)

# 输出关键洞察
print(f"总客户数: {len(rfm):,}")
print(f"高价值客户数: {len(rfm[rfm['Segment'] == '高价值客户']):,}（占比 {len(rfm[rfm['Segment'] == '高价值客户'])/len(rfm):.1%}）")
print("\n=== 高价值客户 Top 10 ===")
print(rfm[rfm["Segment"] == "高价值客户"].nlargest(10, "Monetary")[["Customer ID", "Recency", "Frequency", "Monetary", "RFM_Score"]])

# 保存 RFM 分层图（顺便修复了 palette 警告）
plt.figure(figsize=(10, 6))
# 旧写法（会报警告）：
# sns.countplot(data=rfm, x="Segment", palette="viridis", order=["高价值客户", "中价值客户", "一般客户", "流失风险客户"])
# 新写法（无警告）：
sns.countplot(data=rfm, x="Segment", hue="Segment", palette="viridis", order=["高价值客户", "中价值客户", "一般客户", "流失风险客户"], legend=False)
plt.title("RFM 客户分层分布")
plt.ylabel("客户数量")
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/4_rfm_segments.png", dpi=300, bbox_inches='tight')
plt.close()
print("✅ RFM 分层图已保存")

# ==================== 9. Prophet 下季度销售额预测 ====================

print("\n=== 9. Prophet 下季度销售额预测 ===")

# 准备月度数据（Prophet 推荐）
monthly_prophet = data.resample('MS', on='Order Date')["Sales"].sum().reset_index()
monthly_prophet = monthly_prophet.rename(columns={"Order Date": "ds", "Sales": "y"})

# 建模
model = Prophet(
    yearly_seasonality=True,
    weekly_seasonality=False,
    daily_seasonality=False,
    seasonality_mode='multiplicative'
)
model.fit(monthly_prophet)

# 预测未来 3 个月（下一个完整季度）
future = model.make_future_dataframe(periods=3, freq='MS')
forecast = model.predict(future)

# 打印预测结果
print("=== 下季度销售额预测（USD）===")
next_quarter = forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail(3).round(2)
next_quarter["ds"] = next_quarter["ds"].dt.strftime('%Y-%m')
print(next_quarter)

# 保存预测图
fig1 = model.plot(forecast)
plt.title("全球超市销售额预测（含置信区间） - 下季度预测")
plt.xlabel("时间")
plt.ylabel("销售额 ($)")
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/5_prophet_sales_forecast.png", dpi=300, bbox_inches='tight')
plt.close()

fig2 = model.plot_components(forecast)
plt.savefig(f"{OUTPUT_DIR}/6_prophet_components.png", dpi=300, bbox_inches='tight')
plt.close()

print("✅ Prophet 预测图已保存（总趋势 + 季节性分解）")

print(f"\n🎉 所有分析完成！输出文件夹：{OUTPUT_DIR}")





































































