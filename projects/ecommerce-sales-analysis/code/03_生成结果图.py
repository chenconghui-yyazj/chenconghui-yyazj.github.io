# -*- coding: utf-8 -*-
"""
03_生成结果图.py
重跑电商销售数据全流程（清洗→合并→聚合→RFM→可视化），
输出项目结果图（深色风格，适配简历暗色主题）与关键指标摘要。
"""
import os, json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# ---------- 中文字体 ----------
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# ---------- 深色主题 ----------
BG = "#0b1026"
PANEL = "#0a0e2a"
TXT = "#c9d4ff"
GRID = "#2a3358"
C1 = "#ffb84d"   # 橙
C2 = "#ff4fd8"   # 粉
C3 = "#3b82f6"   # 蓝
C4 = "#7c6bff"   # 紫
C5 = "#39d0a0"   # 青绿
PALETTE = [C3, C1, C2, C5, C4, "#e06c75", "#61afef"]

def style_ax(ax):
    ax.set_facecolor(PANEL)
    ax.figure.set_facecolor(BG)
    for s in ax.spines.values():
        s.set_color(GRID)
    ax.tick_params(colors=TXT, labelsize=10)
    ax.xaxis.label.set_color(TXT)
    ax.yaxis.label.set_color(TXT)
    ax.title.set_color("#ffffff")
    ax.grid(color=GRID, alpha=0.35, linewidth=0.6)
    ax.set_axisbelow(True)

def save(fig, name):
    path = os.path.join(OUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    print("已保存:", path)

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE, "数据集", "电商销售原始数据.xlsx")
OUT_DIR = os.path.join(BASE, "结果图")
os.makedirs(OUT_DIR, exist_ok=True)

# ============================================================
# 读取
# ============================================================
order_detail = pd.read_excel(DATA_PATH, sheet_name="订单明细表")
user_info    = pd.read_excel(DATA_PATH, sheet_name="用户信息表")
product_info = pd.read_excel(DATA_PATH, sheet_name="商品信息表")

od = order_detail.copy(); ui = user_info.copy(); pi = product_info.copy()

# ============================================================
# 清洗（与 Notebook 一致）
# ============================================================
od["数量"] = od["数量"].fillna(od.groupby("品类")["数量"].transform("median"))
mask = od["订单金额"].isna()
od.loc[mask, "订单金额"] = (od.loc[mask, "数量"] * od.loc[mask, "单价"]
                           * (1 - od.loc[mask, "折扣"]) + od.loc[mask, "运费"])

ui["性别"] = ui["性别"].fillna(ui["性别"].mode()[0])
ui["会员等级"] = ui["会员等级"].fillna(ui["会员等级"].mode()[0])
ui["年龄"] = ui["年龄"].fillna(ui["年龄"].median())
pi["品牌"] = pi["品牌"].fillna("未知品牌")
pi["成本价"] = pi["成本价"].fillna(pi.groupby("品类")["成本价"].transform("median"))

od = od.drop_duplicates()
od = od.sort_values("订单金额", ascending=False).drop_duplicates("订单号", keep="first")
od = od[od["数量"] > 0]
od["折扣"] = od["折扣"].clip(0, 0.3)
q1, q3 = od["单价"].quantile(0.25), od["单价"].quantile(0.75)
upper = q3 + 1.5 * (q3 - q1)
od["单价"] = od["单价"].clip(upper=upper)
bad_age = (ui["年龄"] < 18) | (ui["年龄"] > 65)
ui.loc[bad_age, "年龄"] = ui["年龄"].median()

od["订单日期"] = pd.to_datetime(od["订单日期"], errors="coerce")
ui["注册日期"] = pd.to_datetime(ui["注册日期"], errors="coerce")
pi["上架日期"] = pd.to_datetime(pi["上架日期"], errors="coerce")
od = od.dropna(subset=["订单日期"])

od["发货城市"] = od["发货城市"].astype(str).str.strip().str.lower()
ui["城市"] = ui["城市"].astype(str).str.strip().str.lower()
od["订单状态"] = od["订单状态"].astype(str).str.strip()
ui["会员等级"] = ui["会员等级"].astype(str).str.strip()

CITY_MAP = {"guangzhou":"广州","shenzhen":"深圳","beijing":"北京","shanghai":"上海",
            "hangzhou":"杭州","chengdu":"成都","wuhan":"武汉","nanjing":"南京",
            "xian":"西安","chongqing":"重庆","suzhou":"苏州","changsha":"长沙",
            "zhengzhou":"郑州","dongguan":"东莞","foshan":"佛山","tianjin":"天津"}
od["发货城市"] = od["发货城市"].map(CITY_MAP).fillna(od["发货城市"])
ui["城市"] = ui["城市"].map(CITY_MAP).fillna(ui["城市"])
def clean_city(x):
    x = str(x).replace("省", "", 1)
    return x if x in set(CITY_MAP.values()) else "其他"
od["发货城市"] = od["发货城市"].apply(clean_city)
ui["城市"] = ui["城市"].apply(clean_city)
od["订单状态"] = od["订单状态"].str.lower().map(
    {"已完成":"已完成","已发货":"已发货","待发货":"待发货","退款":"退款",
     "refund":"退款","完成":"已完成","退款中":"退款"}).fillna(od["订单状态"])
ui["性别"] = ui["性别"].astype(str).str.strip().str.lower().map(
    {"男":"男","m":"男","male":"男","女":"女","f":"女","female":"女"})
ui["会员等级"] = ui["会员等级"].str.replace(" ", "", regex=False).map(
    {"普通会员":"普通会员","黄金会员":"黄金会员","铂金会员":"铂金会员","钻石会员":"钻石会员"})
ui["会员等级分"] = ui["会员等级"].map({"普通会员":1,"黄金会员":2,"铂金会员":3,"钻石会员":4}).fillna(1)

od = od.reset_index(drop=True); ui = ui.reset_index(drop=True); pi = pi.reset_index(drop=True)

# ============================================================
# 衍生 + 合并
# ============================================================
od["折扣后单价"] = od["单价"] * (1 - od["折扣"])
od["订单月份"] = od["订单日期"].dt.to_period("M").astype(str)
cost_map = pi.set_index("商品ID")["成本价"]
od["成本"] = od["商品ID"].map(cost_map)
od["毛利"] = od["订单金额"] - od["数量"] * od["成本"]

df = od.merge(ui, on="用户ID", how="left").merge(
    pi[["商品ID", "品牌", "上架日期"]], on="商品ID", how="left")

# ============================================================
# 指标
# ============================================================
total_sales = od["订单金额"].sum()
total_orders = od["订单号"].nunique()
total_users = od["用户ID"].nunique()
total_qty = od["数量"].sum()
aov = total_sales / total_orders
item_price = total_sales / total_qty
order_counts = od.groupby("用户ID")["订单号"].nunique()
repurchase = (order_counts >= 2).mean() * 100

by_cat = od.groupby("品类").agg(销售额=("订单金额","sum"), 销量=("数量","sum"),
                               订单数=("订单号","nunique")).sort_values("销售额", ascending=False)
by_cat["客单价"] = (by_cat["销售额"] / by_cat["订单数"]).round(2)
city = od.groupby("发货城市")["订单金额"].sum().sort_values(ascending=False)

monthly = od.set_index("订单日期").resample("ME")["订单金额"].sum()
monthly.index = monthly.index.strftime("%Y-%m")

# RFM
ref = od["订单日期"].max() + pd.Timedelta(days=1)
rfm = od.groupby("用户ID").agg(R=("订单日期", lambda x: (ref - x.max()).days),
                               F=("订单号","nunique"), M=("订单金额","sum"))
rfm["R分"] = (rfm["R"] <= rfm["R"].median()).astype(int)
rfm["F分"] = (rfm["F"] >  rfm["F"].median()).astype(int)
rfm["M分"] = (rfm["M"] >  rfm["M"].median()).astype(int)
def rfm_label(r):
    R,F,M = r["R分"], r["F分"], r["M分"]
    if R and F and M: return "重要价值客户"
    if R and not F and M: return "重要发展客户"
    if not R and F and M: return "重要保持客户"
    if not R and not F and M: return "重要挽留客户"
    if R and F and not M: return "一般价值客户"
    if R and not F and not M: return "一般发展客户"
    if not R and F and not M: return "一般保持客户"
    return "一般挽留客户"
rfm["客户类型"] = rfm.apply(rfm_label, axis=1)
rfm_counts = rfm["客户类型"].value_counts()

# ============================================================
# 图1 月度销售额趋势
# ============================================================
fig, ax = plt.subplots(figsize=(9, 4.2))
ax.plot(monthly.index, monthly.values, marker="o", color=C1, linewidth=2.2, markersize=6)
ax.fill_between(range(len(monthly)), monthly.values, color=C1, alpha=0.12)
ax.set_title("月度销售额趋势（2024）", fontsize=14, fontweight="bold")
ax.set_xlabel("月份"); ax.set_ylabel("销售额（元）")
plt.setp(ax.get_xticklabels(), rotation=40, ha="right")
ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x,_: f"{x/10000:.0f}万"))
style_ax(ax); save(fig, "01_月度销售趋势.png")

# ============================================================
# 图2 品类销售额
# ============================================================
fig, ax = plt.subplots(figsize=(8.5, 4.4))
bars = ax.bar(by_cat.index, by_cat["销售额"], color=PALETTE[:len(by_cat)], width=0.62)
ax.set_title("各品类销售额对比", fontsize=14, fontweight="bold")
ax.set_ylabel("销售额（元）")
for b, v in zip(bars, by_cat["销售额"]):
    ax.text(b.get_x()+b.get_width()/2, v, f"{v/10000:.0f}万", ha="center", va="bottom",
            color=TXT, fontsize=9)
ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x,_: f"{x/10000:.0f}万"))
style_ax(ax); save(fig, "02_品类销售额.png")

# ============================================================
# 图3 品类占比（环形）
# ============================================================
fig, ax = plt.subplots(figsize=(6.2, 5.2))
wedges, texts, autotexts = ax.pie(by_cat["销售额"], labels=by_cat.index, autopct="%.1f%%",
        startangle=90, colors=PALETTE[:len(by_cat)], wedgeprops=dict(width=0.42, edgecolor=BG, linewidth=2),
        textprops=dict(color=TXT, fontsize=9))
for at in autotexts: at.set_color("#ffffff"); at.set_fontsize(9)
ax.set_title("各品类销售额占比", fontsize=14, fontweight="bold", color="#fff")
ax.figure.set_facecolor(BG); save(fig, "03_品类占比.png")

# ============================================================
# 图4 订单金额分布
# ============================================================
fig, ax = plt.subplots(figsize=(8.5, 4.2))
ax.hist(od["订单金额"], bins=50, color=C3, edgecolor=BG, alpha=0.9)
ax.set_title("订单金额分布（右偏长尾）", fontsize=14, fontweight="bold")
ax.set_xlabel("订单金额（元）"); ax.set_ylabel("订单数")
style_ax(ax); save(fig, "04_订单金额分布.png")

# ============================================================
# 图5 品类金额箱线图
# ============================================================
fig, ax = plt.subplots(figsize=(9, 4.8))
data = [od.loc[od["品类"]==c, "订单金额"].values for c in by_cat.index]
bp = ax.boxplot(data, tick_labels=by_cat.index, patch_artist=True, showfliers=False,
                medianprops=dict(color="#ffffff", linewidth=1.6))
for patch, color in zip(bp["boxes"], PALETTE):
    patch.set_facecolor(color); patch.set_alpha(0.55); patch.set_edgecolor(GRID)
ax.set_title("各品类订单金额分布", fontsize=14, fontweight="bold")
ax.set_ylabel("订单金额（元）"); plt.setp(ax.get_xticklabels(), rotation=25)
style_ax(ax); save(fig, "05_品类金额箱线图.png")

# ============================================================
# 图6 相关性热力图
# ============================================================
num = od[["数量","单价","折扣","运费","订单金额"]]
corr = num.corr()
fig, ax = plt.subplots(figsize=(6.4, 5.2))
im = ax.imshow(corr.values, cmap="magma", vmin=-1, vmax=1)
ax.set_xticks(range(len(corr.columns))); ax.set_xticklabels(corr.columns)
ax.set_yticks(range(len(corr.columns))); ax.set_yticklabels(corr.columns)
for i in range(len(corr)):
    for j in range(len(corr)):
        ax.text(j, i, f"{corr.values[i,j]:.2f}", ha="center", va="center",
                color="#ffffff", fontsize=8.5)
ax.set_title("数值指标相关性矩阵", fontsize=14, fontweight="bold")
ax.set_facecolor(BG); ax.figure.set_facecolor(BG)
ax.tick_params(colors=TXT, labelsize=9)
for s in ax.spines.values(): s.set_color(GRID)
save(fig, "06_相关性热力图.png")

# ============================================================
# 图7 RFM 分层
# ============================================================
order = ["重要价值客户","重要发展客户","重要保持客户","重要挽留客户",
         "一般价值客户","一般发展客户","一般保持客户","一般挽留客户"]
vals = [rfm_counts.get(k,0) for k in order]
fig, ax = plt.subplots(figsize=(9.5, 4.6))
bars = ax.barh(order[::-1], vals[::-1], color=(PALETTE*2)[:len(order)])
ax.set_title("RFM 用户分层（人数）", fontsize=14, fontweight="bold")
ax.set_xlabel("用户数")
for b, v in zip(bars, vals[::-1]):
    ax.text(v, b.get_y()+b.get_height()/2, f" {v}", va="center", color=TXT, fontsize=9)
style_ax(ax); save(fig, "07_RFM分层.png")

# ============================================================
# 图8 城市销售额 Top10
# ============================================================
top_city = city.head(10)
fig, ax = plt.subplots(figsize=(8.8, 4.6))
bars = ax.barh(top_city.index[::-1], top_city.values[::-1], color=C5, alpha=0.85)
ax.set_title("城市销售额 Top10", fontsize=14, fontweight="bold")
ax.set_xlabel("销售额（元）")
ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x,_: f"{x/10000:.0f}万"))
style_ax(ax); save(fig, "08_城市销售额.png")

# ============================================================
# 图9 品类×城市热力图
# ============================================================
pv = pd.pivot_table(od, index="品类", columns="发货城市", values="订单金额",
                    aggfunc="sum", fill_value=0)
pv = pv.loc[by_cat.index, city.head(8).index]
fig, ax = plt.subplots(figsize=(9.6, 4.6))
im = ax.imshow(pv.values, cmap="viridis", aspect="auto")
ax.set_xticks(range(pv.shape[1])); ax.set_xticklabels(pv.columns, rotation=35)
ax.set_yticks(range(pv.shape[0])); ax.set_yticklabels(pv.index)
ax.set_title("品类 × 城市 销售额热力图", fontsize=14, fontweight="bold")
ax.set_facecolor(BG); ax.figure.set_facecolor(BG)
ax.tick_params(colors=TXT, labelsize=9)
for s in ax.spines.values(): s.set_color(GRID)
save(fig, "09_品类城市热力图.png")

# ============================================================
# 指标摘要
# ============================================================
summary = {
    "清洗后订单数": int(len(od)),
    "订单总数(去重)": int(total_orders),
    "下单用户数": int(total_users),
    "总销售额": round(float(total_sales), 2),
    "总销量": int(total_qty),
    "客单价": round(float(aov), 2),
    "件单价": round(float(item_price), 2),
    "复购率%": round(float(repurchase), 2),
    "Top品类": by_cat.index[0],
    "Top品类销售额": round(float(by_cat["销售额"].iloc[0]), 2),
    "Top城市": city.index[0],
    "Top城市销售额": round(float(city.iloc[0]), 2),
    "RFM分层": {k: int(rfm_counts.get(k,0)) for k in order},
    "清洗前后行数": {"原始订单": int(len(order_detail)), "清洗后订单": int(len(od))},
}
with open(os.path.join(BASE, "分析摘要.json"), "w", encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print("\n===== 关键指标 =====")
print(f"总销售额 {total_sales:,.0f} 元 | 订单 {total_orders:,} 单 | 用户 {total_users:,} 人")
print(f"客单价 {aov:,.2f} 元 | 件单价 {item_price:,.2f} 元 | 复购率 {repurchase:.2f}%")
print("Top品类:", by_cat.index[0], f"{by_cat['销售额'].iloc[0]:,.0f} 元")
print("Top城市:", city.index[0], f"{city.iloc[0]:,.0f} 元")
print("RFM分层:", summary["RFM分层"])
print("\n品类汇总:\n", by_cat.round(0).to_string())
