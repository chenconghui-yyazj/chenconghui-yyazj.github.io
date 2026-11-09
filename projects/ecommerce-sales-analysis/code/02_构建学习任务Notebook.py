# -*- coding: utf-8 -*-
"""
02_构建学习任务Notebook.py
生成覆盖 Python 数据分析 + 可视化全部知识点的综合实战 Notebook。
产出：02_学习任务与代码.ipynb（可直接用 Jupyter / 豆包 Notebook 打开运行）
"""
import nbformat as nbf
import os

nb = nbf.v4.new_notebook()
nb.metadata = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python", "version": "3.13"},
}
cells = []

def md(text):    cells.append(nbf.v4.new_markdown_cell(text))
def code(text):  cells.append(nbf.v4.new_code_cell(text))

# ============================================================
# 封面与目录
# ============================================================
md("""# 📊 Python 数据分析与可视化 · 综合实战项目（大厂面试标准）

> 项目主题：**电商销售数据分析**（订单明细 / 用户信息 / 商品信息 三张表）
> 学习目标：从 0 到 1 完成「读取 → 清洗 → 整理 → 聚合 → 业务分析 → 可视化 → 结论」的完整链路，
> 覆盖大厂数据分析 / BI 岗面试所需的核心 Python 技能点。

## 📁 项目文件结构
```
电商数据分析综合项目/
├── 数据集/电商销售原始数据.xlsx   ← 原始（脏）数据，供实战清洗
├── 01_生成原始数据集.py            ← 数据生成脚本（可复现）
└── 02_学习任务与代码.ipynb        ← 本文件：全部学习任务与代码
```

## 🗺️ 学习地图（10 个模块 = 全部知识点）
| 模块 | 核心知识点 |
|---|---|
| 0 准备 | 环境、库、项目说明 |
| 1 读取与概览 | `read_excel` / `info` / `describe` / `dtypes` / `shape` / `head` |
| 2 数据清洗 | 缺失值、重复值、异常值、类型转换、字符串清洗、口径统一 |
| 3 整理与合并 | 列操作、筛选、排序、`merge`/`concat`、新增计算列 |
| 4 分组与聚合 | `groupby` / `agg` / `pivot_table` / `crosstab` / 时间重采样 |
| 5 业务指标 | 销售额、客单价、复购率、RFM、帕累托、分层对比 |
| 6 matplotlib | 折线、柱状、饼图、直方图、箱线、散点、子图 |
| 7 seaborn | 分布图、热力图、`pairplot`、分类图、小提琴图 |
| 8 plotly | 交互式折线/柱状/散点/热力图 |
| 9 综合实战 | 完整分析流程 + 大厂面试问答 + 自测题 |
| 10 总结 | 面试高频考点回顾 |

> 💡 **建议用法**：按顺序从上往下逐个运行。每个代码块都带详细中文注释，遇到不理解的先看注释，再动手改参数试。
""")

# ============================================================
# 模块0：准备
# ============================================================
md("""---
## 模块 0 · 环境准备与项目说明

**本模块知识点**：安装与导入第三方库、中文字体设置、项目路径管理。

运行前请在终端确保装好：`pip install pandas numpy matplotlib seaborn plotly openpyxl`
（本项目已经帮你装好了 plotly 和 nbformat，matplotlib/seaborn/pandas 也已就绪。）
""")
code("""# 0.1 导入数据分析所需的全部库
import pandas as pd          # 数据分析核心库：读表、清洗、聚合
import numpy as np           # 科学计算库：数组、随机数、数学运算
import matplotlib.pyplot as plt   # 基础绘图库（静态图）
import seaborn as sns        # 高级统计绘图库（基于 matplotlib）
import plotly.express as px  # 交互式绘图库（可缩放/悬浮）
import os

# 0.2 设置 matplotlib 中文字体（Windows 用微软雅黑；避免中文显示成方块）
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False   # 解决负号显示成方块的问题

# 0.3 用绝对路径定位原始数据文件（请改成你自己的电脑路径）
DATA_PATH = r'C:\\Users\\陈聪惠\\Doubao\\chats\\2026-11-05\\new-chat-1\\电商数据分析综合项目\\数据集\\电商销售原始数据.xlsx'
print('数据文件是否存在：', os.path.exists(DATA_PATH))
""")

# ============================================================
# 模块1：读取与概览
# ============================================================
md("""---
## 模块 1 · 数据读取与概览

**本模块知识点**：用 pandas 读取 Excel 的多个子表；用 `head/tail/sample/info/describe` 快速了解数据规模、字段类型、缺失情况与数值分布。

> 面试考察点：拿到一份数据，第一步是"摸清结构"——多少行多少列、每列是什么类型、有没有缺失、数值分布大概什么样。
""")
code("""# 1.1 一次性读取 Excel 中的三张子表
order_detail = pd.read_excel(DATA_PATH, sheet_name='订单明细表')   # 每一笔订单的销售记录
user_info    = pd.read_excel(DATA_PATH, sheet_name='用户信息表')   # 用户档案
product_info = pd.read_excel(DATA_PATH, sheet_name='商品信息表')   # 商品档案

# 1.2 查看每张表的规模（shape 返回 (行数, 列数)）
print('订单明细表:', order_detail.shape)
print('用户信息表:', user_info.shape)
print('商品信息表:', product_info.shape)

# 1.3 看前几行 / 随机几行，直观感受数据内容
print('\\n---- 订单明细表前 3 行 ----')
print(order_detail.head(3))
print('\\n---- 随机抽 2 个用户 ----')
print(user_info.sample(2, random_state=1))
""")
code("""# 1.4 info()：快速总览每列类型与非空数量（能立刻看出哪些列有缺失）
print('===== 订单明细表结构 =====')
order_detail.info()
print('\\n===== 用户信息表结构 =====')
user_info.info()
""")
code("""# 1.5 describe()：数值列的统计描述（count/mean/std/min/四分位/max）
#     一眼看出异常：比如"数量"的 min 可能为负、"年龄"的 max 可能离谱
print('===== 订单明细表数值统计 =====')
print(order_detail.describe().round(2))
print('\\n===== 用户信息表数值统计 =====')
print(user_info.describe().round(2))
""")
code("""# 1.6 dtypes：查看每列的数据类型（object=文本/混合，int/float=数值，datetime=日期）
print('订单明细表各列类型：')
print(order_detail.dtypes)

# 1.7 nunique()：每列有多少种不同取值（文本列尤其要看，方便发现口径不统一）
print('\\n各列不同取值个数：')
print(order_detail.nunique())
print('\\n性别列取值：', user_info['性别'].unique())
print('订单状态列取值：', order_detail['订单状态'].unique())
print('发货城市列取值(前10)：', order_detail['发货城市'].unique()[:10])
""")

# ============================================================
# 模块2：数据清洗
# ============================================================
md("""---
## 模块 2 · 数据清洗（整个项目的重点与核心面试考点）

**本模块知识点**：
- **缺失值**：`isnull/isna` 检测、缺失占比、`fillna` 填充、`dropna` 删除
- **重复值**：`duplicated` 检测、`drop_duplicates` 去重（全字段重复 vs 主键重复）
- **异常值**：业务规则 + 箱线图 + IQR（四分位距）法 + 3σ 法则
- **类型转换**：`to_datetime`、`to_numeric`、`astype`（处理混合格式日期）
- **字符串清洗**：`strip` 去空格、`lower/upper` 统一大小写、`replace` 映射统一口径
- **自定义函数**：`apply` / `map` / `applymap`

> 面试考察点：数据清洗最能体现分析师的"基本功"。原则是**先诊断，后处理**——每一步都要说明为什么这么处理。
""")
md("### 2.1 缺失值处理""")
code("""# 2.1.1 诊断缺失：isnull()/isna() 返回 True/False 布尔矩阵，sum() 统计每列缺失个数
print('===== 各表缺失情况 =====')
print('订单明细表缺失数：')
print(order_detail.isna().sum())
print('\\n用户信息表缺失数：')
print(user_info.isna().sum())
print('\\n商品信息表缺失数：')
print(product_info.isna().sum())

# 缺失占比：缺失数 / 总行数
print('\\n订单明细表"数量"缺失占比：', round(order_detail['数量'].isna().mean()*100, 2), '%')
""")
code("""# 2.1.2 制定并执行缺失值处理策略
# 先复制原始表，避免污染源数据（好的习惯）
od = order_detail.copy()   # od = 清洗后的订单明细
ui = user_info.copy()      # ui = 清洗后的用户信息
pi = product_info.copy()   # pi = 清洗后的商品信息

# —— 策略1：'数量'缺失 -> 用该品类的"数量中位数"填充（fillna + groupby 变换）——
# groupby('品类')['数量'].transform('median')：给每一行算出它所属品类的中位数
od['数量'] = od['数量'].fillna(od.groupby('品类')['数量'].transform('median'))
print('数量缺失填充后剩余缺失：', od['数量'].isna().sum())

# —— 策略2：'订单金额'缺失 -> 用公式重算填充（数量*单价*(1-折扣)+运费）——
mask = od['订单金额'].isna()
od.loc[mask, '订单金额'] = (od.loc[mask, '数量'] * od.loc[mask, '单价']
                            * (1 - od.loc[mask, '折扣']) + od.loc[mask, '运费'])
print('订单金额缺失重算后剩余缺失：', od['订单金额'].isna().sum())
""")
code("""# 2.1.3 用户表缺失处理
# '性别'缺失：由于无法凭空推断，这里用"众数(出现最多的值)"填充
ui['性别'] = ui['性别'].fillna(ui['性别'].mode()[0])

# '会员等级'缺失：用众数填充
ui['会员等级'] = ui['会员等级'].fillna(ui['会员等级'].mode()[0])

# '年龄'缺失：先看分布，再用中位数填充（中位数对异常值更稳健）
print('年龄缺失前中位数：', ui['年龄'].median())
ui['年龄'] = ui['年龄'].fillna(ui['年龄'].median())

# 商品表 '品牌'/'成本价' 缺失：品牌用'未知品牌'，成本价用该品类中位数
pi['品牌'] = pi['品牌'].fillna('未知品牌')
pi['成本价'] = pi['成本价'].fillna(pi.groupby('品类')['成本价'].transform('median'))
print('清洗后各表剩余缺失合计：',
      od.isna().sum().sum(), ui.isna().sum().sum(), pi.isna().sum().sum())
""")
md("### 2.2 重复值处理""")
code("""# 2.2.1 诊断重复
# duplicated() 默认按"所有列完全相同"判断，返回布尔序列
print('全字段完全重复的行数：', od.duplicated().sum())

# 按"订单号"看重复（一个订单号本应只出现一次）
print('订单号出现重复的次数：', od['订单号'].duplicated().sum())

# 看具体的重复订单号长什么样
dup_oid = od['订单号'][od['订单号'].duplicated()].unique()
print('\\n示例重复订单号：', dup_oid[:5])
print(od[od['订单号'].isin(dup_oid)].head(6))
""")
code("""# 2.2.2 去重策略
# 第1步：删除"全字段完全重复"的行（数据录入重复）
od = od.drop_duplicates()
print('删除全字段重复后行数：', len(od))

# 第2步：同一订单号若出现多行且内容不同，属于"异常重复"——保留金额最大的一条，删掉其余
od = od.sort_values('订单金额', ascending=False).drop_duplicates('订单号', keep='first')
od = od.reset_index(drop=True)   # 重新生成连续行索引
print('按订单号去重后行数：', len(od))
print('最终订单表规模：', od.shape)
""")
md("### 2.3 异常值处理""")
code("""# 2.3.1 用业务规则诊断异常
# 数量不能为负或 0；折扣一般在 0~0.3；年龄应在 18~65 之间
print('数量 <= 0 的行数：', (od['数量'] <= 0).sum())
print('折扣 > 0.3 的行数：', (od['折扣'] > 0.3).sum())
print('折扣 < 0 的行数：', (od['折扣'] < 0).sum())
print('\\n年龄异常(不在18~65)的用户数：', ((ui['年龄'] < 18) | (ui['年龄'] > 65)).sum())
print('异常年龄示例：', ui.loc[(ui['年龄'] < 18) | (ui['年龄'] > 65), '年龄'].tolist()[:10])
""")
code("""# 2.3.2 IQR（四分位距）法检测单价异常值
# IQR = Q3 - Q1；一般认为 < Q1-1.5*IQR 或 > Q3+1.5*IQR 为离群值
q1 = od['单价'].quantile(0.25)
q3 = od['单价'].quantile(0.75)
iqr = q3 - q1
lower, upper = q1 - 1.5*iqr, q3 + 1.5*iqr
print(f'单价：Q1={q1:.2f}, Q3={q3:.2f}, IQR={iqr:.2f}')
print(f'正常范围：[{lower:.2f}, {upper:.2f}]')
print('超出上界的行数：', (od['单价'] > upper).sum())
""")
code("""# 2.3.3 执行异常值清洗
# 数量：<=0 的视为录入错误，删除
od = od[od['数量'] > 0]
# 折扣：>0.3 视为错误，裁剪到 0.3；<0 置 0
od['折扣'] = od['折扣'].clip(0, 0.3)
# 单价：超过上界的用上界封顶（winsorize 截尾法，比直接删更稳妥）
od['单价'] = od['单价'].clip(upper=upper)
# 年龄：超出 18~65 的视为无效，用中位数填充
bad_age = (ui['年龄'] < 18) | (ui['年龄'] > 65)
ui.loc[bad_age, '年龄'] = ui['年龄'].median()
print('清洗后订单表规模：', od.shape)
print('清洗后年龄范围：', ui['年龄'].min(), '~', ui['年龄'].max())
""")
md("### 2.4 数据类型转换（处理混合格式日期）""")
code("""# 2.4.1 订单日期是"混合格式"（有的是 datetime，有的是 '2024/01/15' 字符串，有的是 '20240115'）
# pd.to_datetime 能智能解析多种格式；errors='coerce' 表示解析失败变成 NaT(缺失)
od['订单日期'] = pd.to_datetime(od['订单日期'], errors='coerce')
ui['注册日期'] = pd.to_datetime(ui['注册日期'], errors='coerce')
pi['上架日期'] = pd.to_datetime(pi['上架日期'], errors='coerce')
print('订单日期类型：', od['订单日期'].dtype)
print('解析失败(变成NaT)的行数：', od['订单日期'].isna().sum())

# 2.4.2 顺带把数值列统一为数值类型（to_numeric 兜底，遇到非数字变 NaN）
od['数量'] = pd.to_numeric(od['数量'], errors='coerce')
od['单价'] = pd.to_numeric(od['单价'], errors='coerce')
od['订单金额'] = pd.to_numeric(od['订单金额'], errors='coerce')

# 2.4.3 astype：把某些列转成明确的类型
# 例如把"用户ID"转成整数
od['用户ID'] = od['用户ID'].astype('int64')
ui['用户ID'] = ui['用户ID'].astype('int64')
print('\\n清洗后订单明细表字段类型：')
print(od.dtypes)
""")
md("### 2.5 字符串清洗与口径统一""")
code("""# 2.5.1 去除首尾空格（strip）+ 统一大小写（lower）
# 城市里混着 ' 广州'、'Guangzhou'、'shenzhen' 等
od['发货城市'] = od['发货城市'].astype(str).str.strip().str.lower()
ui['城市'] = ui['城市'].astype(str).str.strip().str.lower()
od['订单状态'] = od['订单状态'].astype(str).str.strip()
ui['会员等级'] = ui['会员等级'].astype(str).str.strip()

print('城市取值：', od['发货城市'].unique()[:15])
print('状态取值：', od['订单状态'].unique())
""")
code("""# 2.5.2 用映射(map/dict)统一口径：把大小写英文统一成标准中文
CITY_MAP = {'guangzhou': '广州', 'shenzhen': '深圳', 'beijing': '北京', 'shanghai': '上海',
            'hangzhou': '杭州', 'chengdu': '成都', 'wuhan': '武汉', 'nanjing': '南京',
            'xian': '西安', 'chongqing': '重庆', 'suzhou': '苏州', 'changsha': '长沙',
            'zhengzhou': '郑州', 'dongguan': '东莞', 'foshan': '佛山', 'tianjin': '天津'}
# 英文名通过映射统一成中文；中文城市名（东莞/西安…）原样保留，不要被归到'其他'
od['发货城市'] = od['发货城市'].map(CITY_MAP).fillna(od['发货城市'])
ui['城市'] = ui['城市'].map(CITY_MAP).fillna(ui['城市'])
# 剔除脏值：去掉'省'后缀（如'深圳省'），并把不在已知城市名单里的归为'其他'
def clean_city(x):
    x = str(x).replace('省', '', 1)
    return x if x in set(CITY_MAP.values()) else '其他'
od['发货城市'] = od['发货城市'].apply(clean_city)
ui['城市'] = ui['城市'].apply(clean_city)
print('统一后的发货城市：', od['发货城市'].unique())

# 状态映射：把 'REFUND'、'完成' 等统一为规范状态
STATUS_MAP = {'已完成': '已完成', '已发货': '已发货', '待发货': '待发货', '退款': '退款',
              'refund': '退款', '完成': '已完成', '已完成 ': '已完成', ' 已发货': '已发货',
              '退款中': '退款'}
od['订单状态'] = od['订单状态'].str.lower().map(
    {'已完成': '已完成', '已发货': '已发货', '待发货': '待发货', '退款': '退款',
     'refund': '退款', '完成': '已完成', '退款中': '退款'}).fillna(od['订单状态'])
print('统一后的状态：', od['订单状态'].unique())
""")
code("""# 2.5.3 性别统一（男/M/male -> 男；女/F/female -> 女）
GENDER_MAP = {'男': '男', 'M': '男', 'male': '男', '男 ': '男',
              '女': '女', 'F': '女', 'female': '女', ' 女': '女'}
ui['性别'] = ui['性别'].astype(str).str.strip().str.lower().map(
    {'男': '男', 'm': '男', 'male': '男', '女': '女', 'f': '女', 'female': '女'})
# 会员等级去空格已经做过，再做一次规范映射（含空格变体）
LEVEL_MAP = {'普通会员': '普通会员', '黄金会员': '黄金会员', '铂金会员': '铂金会员', '钻石会员': '钻石会员'}
ui['会员等级'] = ui['会员等级'].str.replace(' ', '', regex=False).map(LEVEL_MAP)
print('性别取值：', ui['性别'].unique())
print('会员等级取值：', ui['会员等级'].unique())

# 2.5.4 apply 自定义函数：把会员等级映射成数字（方便后续排序/分层）
def level_to_score(x):
    return {'普通会员': 1, '黄金会员': 2, '铂金会员': 3, '钻石会员': 4}.get(x, 1)
ui['会员等级分'] = ui['会员等级'].apply(level_to_score)
print('会员等级分示例：', ui[['会员等级', '会员等级分']].head(3).to_dict('records'))

# 2.5.5 清洗收尾：重置索引，打印最终规模
od = od.reset_index(drop=True)
ui = ui.reset_index(drop=True)
pi = pi.reset_index(drop=True)
print('\\n✅ 数据清洗完成！')
print('订单明细表：', od.shape, ' 用户信息表：', ui.shape, ' 商品信息表：', pi.shape)
""")

# ============================================================
# 模块3：整理与合并
# ============================================================
md("""---
## 模块 3 · 数据整理与合并

**本模块知识点**：列的选择/新增/重命名、布尔条件筛选（`query`/`isin`/`between`）、排序 `sort_values`、多表合并 `merge`/`concat`、计算列。

> 面试考察点：多表 join 是分析的基础，要理解内连接/左连接的区别，以及"一对一/一对多"合并。
""")
code("""# 3.1 列操作：选择、新增、重命名
# 选择指定列
subset = od[['订单号', '用户ID', '订单日期', '品类', '数量', '订单金额']]
print(subset.head(3))

# 新增计算列：折扣后单价 = 单价 * (1-折扣)；订单月份；毛利润 = 金额 - 成本
od['折扣后单价'] = od['单价'] * (1 - od['折扣'])
od['订单月份'] = od['订单日期'].dt.to_period('M').astype(str)   # 例如 '2024-01'
# 合并商品成本，算毛利（merge 见 3.4）
cost_map = pi.set_index('商品ID')['成本价']
od['成本'] = od['商品ID'].map(cost_map)
od['毛利'] = od['订单金额'] - od['数量'] * od['成本']
print('\\n新增列后字段：', list(od.columns))
print(od[['折扣后单价', '订单月份', '成本', '毛利']].head(3))
""")
code("""# 3.2 条件筛选：布尔索引 / query / isin / between
# 筛选"数码类且金额>=500"的订单
mask = (od['品类'] == '数码') & (od['订单金额'] >= 500)
print('数码高客单订单数：', mask.sum())
print(od.loc[mask, ['订单号', '商品名称', '订单金额']].head(5))

# query() 写法（更接近 SQL，代码更清爽）
q = od.query('品类 == "数码" and 订单金额 >= 500')
print('\\nquery 结果条数：', len(q))

# isin：判断是否属于某个集合
top_cities = ['广州', '深圳', '北京']
print('\\n一线/新一线大城订单数：', od['发货城市'].isin(top_cities).sum())

# between：数值区间判断
print('单价在100~500的订单数：', od['单价'].between(100, 500).sum())
""")
code("""# 3.3 排序 sort_values
# 按订单金额降序（大到小），看 top10 大单
top10 = od.sort_values('订单金额', ascending=False).head(10)
print('金额 Top10 订单：')
print(top10[['订单号', '商品名称', '数量', '单价', '订单金额', '发货城市']])

# 多列排序：先按月份，再按金额降序
od_sorted = od.sort_values(['订单月份', '订单金额'], ascending=[True, False])
print('\\n按月份+金额排序后前3行：')
print(od_sorted[['订单月份', '订单金额']].head(3))
""")
code("""# 3.4 多表合并 merge / concat
# merge 内连接：订单表按 用户ID 关联 用户表，给每笔订单补上用户信息
df = od.merge(ui, on='用户ID', how='left')   # how='left' 左连接：保留全部订单
print('订单+用户合并后规模：', df.shape)
print('新增字段：', list(df.columns)[-6:])

# 再关联商品表，补上品类品牌信息（其实订单表已有品类，这里演示用商品ID关联成本等）
df = df.merge(pi[['商品ID', '品牌', '上架日期']], on='商品ID', how='left')
print('合并全部后规模：', df.shape)

# concat：纵向拼接（常用于追加新数据），这里演示把两个城市子集拼回来
g = od[od['发货城市'] == '广州'].copy()
s = od[od['发货城市'] == '深圳'].copy()
con = pd.concat([g, s], ignore_index=True)
print('\\nconcat 拼接广州+深圳订单行数：', len(con), '=', len(g)+len(s))
""")

# ============================================================
# 模块4：分组与聚合
# ============================================================
md("""---
## 模块 4 · 分组与聚合分析

**本模块知识点**：`groupby` 单列/多列分组、`agg` 多指标聚合、`pivot_table` 透视表、`crosstab` 交叉表、时间序列 `resample` 重采样与环比/同比。

> 面试考察点：分组聚合是 SQL 与 pandas 共同的核心，"group by 能完成什么"必须脱口而出。
""")
code("""# 4.1 groupby 单列分组：按品类统计销售额、销量、订单数
by_cat = od.groupby('品类').agg(
    销售额=('订单金额', 'sum'),
    销量=('数量', 'sum'),
    订单数=('订单号', 'nunique'),
    客单价=('订单金额', 'mean'),
).round(2).sort_values('销售额', ascending=False)
print(by_cat)
""")
code("""# 4.2 groupby 多列分组：按"品类 + 订单状态"
by_cat_status = od.groupby(['品类', '订单状态'])['订单金额'].sum().unstack(fill_value=0)
print(by_cat_status.round(0))
""")
code("""# 4.3 pivot_table 透视表：行列=维度，值=指标，aggfunc=聚合方式
pt = pd.pivot_table(od, index='品类', columns='订单状态',
                    values='订单金额', aggfunc='sum', fill_value=0).round(0)
print('透视表(品类×状态 销售额)：')
print(pt)
""")
code("""# 4.4 crosstab 交叉表：统计"会员等级 × 性别"的人数
ct = pd.crosstab(ui['会员等级'], ui['性别'])
print('交叉表(会员等级×性别 人数)：')
print(ct)

# 加 normalize 看占比
ct_pct = pd.crosstab(ui['会员等级'], ui['性别'], normalize='index').round(3)
print('\\n按行归一化(每行和为1)：')
print(ct_pct)
""")
code("""# 4.5 时间序列分析：按月聚合销售额
od['订单日期'] = pd.to_datetime(od['订单日期'])
monthly = od.set_index('订单日期').resample('ME')['订单金额'].sum()  # ME=每月末
monthly = monthly.reset_index()
monthly['月份'] = monthly['订单日期'].dt.to_period('M').astype(str)
print('月度销售额：')
print(monthly[['月份', '订单金额']].round(2).to_string(index=False))

# 环比：本月 / 上月 - 1
monthly['环比'] = monthly['订单金额'].pct_change() * 100
# 同比：去年同期（这里只有一年数据，演示用第一个月做基准）
monthly['累计'] = monthly['订单金额'].cumsum()
print('\\n含环比/累计：')
print(monthly[['月份', '订单金额', '环比', '累计']].round(2).to_string(index=False))
""")

# ============================================================
# 模块5：业务指标分析
# ============================================================
md("""---
## 模块 5 · 业务指标分析（大厂面试最看重的部分）

**本模块知识点**：总销售额/订单量/客单价/复购率等基础指标；城市、品类、会员分层对比；**RFM 用户分层**；**帕累托(二八)法则**；漏斗与留存思路。

> 面试考察点：不只算数，要能说出"指标的业务含义"和"指导什么决策"。例如客单价低该推什么、复购率低该做什么。
""")
code("""# 5.1 核心经营指标
total_sales = od['订单金额'].sum()              # 总销售额
total_orders = od['订单号'].nunique()           # 订单数(去重)
total_users = od['用户ID'].nunique()            # 下单用户数
total_qty = od['数量'].sum()                    # 总销量
avg_price = total_sales / total_orders          # 客单价(每笔订单平均金额)
item_price = total_sales / total_qty            # 件单价(每件商品平均价格)

# 复购率：下单次数 >= 2 的用户占比
order_counts = od.groupby('用户ID')['订单号'].nunique()
repurchase_rate = (order_counts >= 2).mean() * 100

print('===== 核心经营指标 =====')
print(f'总销售额：{total_sales:,.0f} 元')
print(f'总订单数：{total_orders:,} 单')
print(f'下单用户数：{total_users:,} 人')
print(f'总销量：{total_qty:,.0f} 件')
print(f'客单价：{avg_price:,.2f} 元/单')
print(f'件单价：{item_price:,.2f} 元/件')
print(f'复购率：{repurchase_rate:.2f}%')
""")
code("""# 5.2 城市维度对比：销售额与客单价
city_stat = od.groupby('发货城市').agg(
    销售额=('订单金额', 'sum'), 订单数=('订单号', 'nunique')
).sort_values('销售额', ascending=False)
city_stat['客单价'] = (city_stat['销售额'] / city_stat['订单数']).round(2)
print('城市经营表现 Top8：')
print(city_stat.head(8).round(0))

# 5.3 会员分层：不同等级用户的消费力
member_stat = df.groupby('会员等级').agg(
    用户数=('用户ID', 'nunique'), 销售额=('订单金额', 'sum')
).sort_values('销售额', ascending=False)
member_stat['人均消费'] = (member_stat['销售额'] / member_stat['用户数']).round(2)
print('\\n会员等级消费力：')
print(member_stat.round(0))
""")
code("""# 5.4 RFM 用户分层（大厂经典模型）
# R=最近一次下单距今多久(Recency, 越小越活跃)、F=下单频次(Frequency)、M=消费金额(Monetary)
import datetime as dt
ref_date = od['订单日期'].max() + pd.Timedelta(days=1)   # 参考日=最后下单日+1

rfm = od.groupby('用户ID').agg(
    R=('订单日期', lambda x: (ref_date - x.max()).days),   # 最近一次下单距参考日的天数
    F=('订单号', 'nunique'),                                # 下单次数
    M=('订单金额', 'sum'),                                  # 累计消费金额
)
# 用中位数切分：高于中位数记 1，否则记 0
rfm['R分'] = (rfm['R'] <= rfm['R'].median()).astype(int)
rfm['F分'] = (rfm['F'] > rfm['F'].median()).astype(int)
rfm['M分'] = (rfm['M'] > rfm['M'].median()).astype(int)

# 组合成 8 种客户类型
def rfm_label(r):
    r, f, m = r['R分'], r['F分'], r['M分']
    if r and f and m:   return '重要价值客户'
    if r and not f and m: return '重要发展客户'
    if not r and f and m: return '重要保持客户'
    if not r and not f and m: return '重要挽留客户'
    if r and f and not m: return '一般价值客户'
    if r and not f and not m: return '一般发展客户'
    if not r and f and not m: return '一般保持客户'
    return '一般挽留客户'
rfm['客户类型'] = rfm.apply(rfm_label, axis=1)

print('RFM 客户分层结果：')
print(rfm['客户类型'].value_counts())
print('\\n各类型平均消费：')
print(rfm.groupby('客户类型')['M'].mean().round(0).sort_values(ascending=False))
""")
code("""# 5.5 帕累托(二八)法则：品类销售额累计占比
cat_sales = od.groupby('品类')['订单金额'].sum().sort_values(ascending=False)
cum_pct = (cat_sales.cumsum() / cat_sales.sum() * 100).round(1)
print('品类销售额累计占比：')
for c, v in cum_pct.items():
    print(f'  {c}: 累计占比 {v}%')
""")

# ============================================================
# 模块6：matplotlib
# ============================================================
md("""---
## 模块 6 · matplotlib 数据可视化（静态图）

**本模块知识点**：折线图(趋势)、柱状图(对比)、饼图(占比)、直方图(分布)、箱线图(异常/离散)、散点图(相关)、子图与样式设置。

> 面试考察点：会看图、会挑图——趋势用折线、对比用柱状、占比用饼图、分布用直方图/箱线、相关用散点。图要"会说话"：标题、轴标签、图例齐全。
""")
code("""# 6.0 准备绘图用的聚合数据
# 月度销售额(折线图用)
m = od.set_index('订单日期').resample('ME')['订单金额'].sum()
m.index = m.index.strftime('%Y-%m')
# 品类销售额(柱状/饼图用)
cat = od.groupby('品类')['订单金额'].sum().sort_values(ascending=False)
# 单价(直方图用)
price = od['单价']
# 客单价按品类(箱线图用)
box_data = [od.loc[od['品类'] == c, '订单金额'] for c in cat.index]
""")
code("""# 6.1 折线图：月度销售额趋势
plt.figure(figsize=(10, 5))
plt.plot(m.index, m.values, marker='o', color='#2E86AB', linewidth=2)
plt.title('2024 年各月销售额趋势', fontsize=14)
plt.xlabel('月份'); plt.ylabel('销售额(元)')
plt.xticks(rotation=45)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.show()
""")
code("""# 6.2 柱状图：各品类销售额对比（横向柱状更易读）
plt.figure(figsize=(9, 5))
plt.barh(cat.index, cat.values, color='#4C9F70')
plt.title('各品类销售额对比', fontsize=14)
plt.xlabel('销售额(元)'); plt.ylabel('品类')
for i, v in enumerate(cat.values):
    plt.text(v, i, f'{v:,.0f}', va='center', fontsize=9)
plt.tight_layout()
plt.show()
""")
code("""# 6.3 饼图：品类销售额占比
plt.figure(figsize=(7, 7))
plt.pie(cat.values, labels=cat.index, autopct='%.1f%%',
        startangle=90, wedgeprops={'edgecolor': 'white'})
plt.title('各品类销售额占比', fontsize=14)
plt.axis('equal')
plt.show()
""")
code("""# 6.4 直方图：单价分布（观察数据形态/是否有长尾）
plt.figure(figsize=(9, 5))
plt.hist(price, bins=50, color='#D98324', edgecolor='white', alpha=0.85)
plt.title('商品单价分布', fontsize=14)
plt.xlabel('单价(元)'); plt.ylabel('订单数')
plt.tight_layout()
plt.show()
""")
code("""# 6.5 箱线图：各品类订单金额分布（看中位数/离群值）
# 注意：matplotlib 3.9+ 把 labels 参数改名为 tick_labels
plt.figure(figsize=(9, 6))
plt.boxplot(box_data, tick_labels=cat.index, showfliers=True)
plt.title('各品类订单金额分布', fontsize=14)
plt.ylabel('订单金额(元)'); plt.xticks(rotation=30)
plt.tight_layout()
plt.show()
""")
code("""# 6.6 散点图：数量 vs 金额（看相关性）
plt.figure(figsize=(8, 5))
plt.scatter(od['数量'], od['订单金额'], alpha=0.4, s=20, color='#5B4B8A')
plt.title('数量与订单金额关系', fontsize=14)
plt.xlabel('数量(件)'); plt.ylabel('订单金额(元)')
plt.tight_layout()
plt.show()
""")
code("""# 6.7 子图：一页展示 4 张图（subplots 布局）
fig, axes = plt.subplots(2, 2, figsize=(13, 9))

# 左上：月度趋势
axes[0,0].plot(m.index, m.values, color='#2E86AB', marker='o')
axes[0,0].set_title('月度销售额'); axes[0,0].tick_params(axis='x', rotation=45)

# 右上：品类占比(条形)
axes[0,1].barh(cat.index, cat.values, color='#4C9F70')
axes[0,1].set_title('品类销售额')

# 左下：单价直方图
axes[1,0].hist(price, bins=40, color='#D98324', edgecolor='white')
axes[1,0].set_title('单价分布')

# 右下：数量-金额散点
axes[1,1].scatter(od['数量'], od['订单金额'], alpha=0.4, color='#5B4B8A', s=15)
axes[1,1].set_title('数量 vs 金额')

plt.tight_layout()
plt.show()
""")

# ============================================================
# 模块7：seaborn
# ============================================================
md("""---
## 模块 7 · seaborn 统计可视化

**本模块知识点**：`histplot/kdeplot` 分布、`countplot/barplot` 分类、`boxplot/violinplot` 箱线/小提琴、`heatmap` 热力图、`pairplot` 多变量矩阵。

> seaborn 基于 matplotlib，但自带好看的默认主题和统计功能（如置信区间），是"分析型图表"的首选。
""")
code("""# 7.0 使用 seaborn 默认主题
sns.set_theme(style='whitegrid', font='Microsoft YaHei')
# 7.1 分布图：直方图+核密度曲线(kde)
plt.figure(figsize=(9, 5))
sns.histplot(od['单价'], bins=60, kde=True, color='#6A9FB5')
plt.title('单价分布（直方图+KDE）', fontsize=14)
plt.show()
""")
code("""# 7.2 分类柱状图 barplot / countplot
plt.figure(figsize=(9, 5))
sns.barplot(x=cat.index, y=cat.values, palette='viridis')
plt.title('各品类销售额', fontsize=14)
plt.ylabel('销售额(元)'); plt.xlabel('品类')
plt.tight_layout()
plt.show()
""")
code("""# 7.3 箱线图 + 小提琴图：对比不同品类的订单金额分布
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
sns.boxplot(data=od, x='品类', y='订单金额', ax=axes[0], palette='Set2')
axes[0].set_title('品类订单金额箱线图'); axes[0].tick_params(axis='x', rotation=30)
sns.violinplot(data=od, x='品类', y='订单金额', ax=axes[1], palette='Set3')
axes[1].set_title('品类订单金额小提琴图'); axes[1].tick_params(axis='x', rotation=30)
plt.tight_layout()
plt.show()
""")
code("""# 7.4 热力图 heatmap：各品类×各城市 销售额
pv = pd.pivot_table(od, index='品类', columns='发货城市',
                    values='订单金额', aggfunc='sum', fill_value=0)
plt.figure(figsize=(12, 5))
sns.heatmap(pv, annot=True, fmt='.0f', cmap='YlGnBu', linewidths=0.5)
plt.title('品类 × 城市 销售额热力图（元）', fontsize=14)
plt.tight_layout()
plt.show()
""")
code("""# 7.5 相关性热力图：看数值变量之间相关性
num_cols = od[['数量', '单价', '折扣', '运费', '订单金额']]
corr = num_cols.corr()
plt.figure(figsize=(7, 6))
sns.heatmap(corr, annot=True, cmap='coolwarm', center=0, fmt='.2f')
plt.title('数值变量相关性矩阵', fontsize=14)
plt.tight_layout()
plt.show()
""")
code("""# 7.6 pairplot 多变量矩阵（数据多时更直观；这里抽样 500 行避免卡顿）
sample = od[['数量', '单价', '订单金额', '折扣']].sample(500, random_state=1)
sns.pairplot(sample, diag_kind='kde', corner=True)
plt.show()
""")

# ============================================================
# 模块8：plotly
# ============================================================
md("""---
## 模块 8 · plotly 交互式可视化

**本模块知识点**：`plotly.express` 一行生成交互图（可缩放、悬浮显示数值、图例点击隐藏），适合做汇报/BI 看板。

> 面试考察点：会用 plotly/Tableau 把静态结论变成"可交互的汇报看板"，是大厂 BI 岗加分项。
""")
code("""# 8.1 交互式折线图：月度销售额
import plotly.express as px
m2 = monthly.copy()   # monthly 来自模块4，含"月份"和"订单金额"两列
fig = px.line(m2, x='月份', y='订单金额', title='2024 月度销售额趋势(交互)',
              markers=True, labels={'订单金额': '销售额(元)'})
fig.show()
""")
code("""# 8.2 交互式柱状图：品类销售额
cat_df = cat.reset_index().rename(columns={'index': '品类', '订单金额': '销售额'})
fig = px.bar(cat_df, x='品类', y='销售额', title='各品类销售额(交互)',
             text_auto='.2s', color='品类')
fig.show()
""")
code("""# 8.3 交互式散点图：数量 vs 金额（颜色=品类）
scatter_df = od.sample(800, random_state=1)
fig = px.scatter(scatter_df, x='数量', y='订单金额', color='品类',
                 hover_data=['商品名称'], title='数量-金额散点(交互)',
                 labels={'数量': '数量(件)', '订单金额': '金额(元)'})
fig.show()
""")
code("""# 8.4 交互式热力图：品类×城市 销售额
pv_r = pv.reset_index().melt(id_vars='品类', var_name='城市', value_name='销售额')
fig = px.density_heatmap(pv_r, x='品类', y='城市', z='销售额',
                         title='品类×城市销售额热力图(交互)', color_continuous_scale='Blues')
fig.show()
""")

# ============================================================
# 模块9：综合实战
# ============================================================
md("""---
## 模块 9 · 综合实战（完整流程）+ 大厂面试问答

**目标**：把前面所有知识点串成一条流水线，一次性输出一份"电商经营分析结论"。

> 面试考察点：能否把一个开放问题拆解为"指标定义→数据处理→分析→图表→结论与建议"的完整闭环。
""")
code("""# 9.1 定义一个"一键分析"函数，串起 清洗→聚合→输出
def analyze_sales(od):
    '''输入清洗后的订单表，输出核心经营结论(结构化)'''
    # 总体
    total = od['订单金额'].sum()
    orders = od['订单号'].nunique()
    aov = total / orders
    # 品类表现
    cat = od.groupby('品类').agg(销售额=('订单金额','sum'),
                                 订单数=('订单号','nunique')).sort_values('销售额', ascending=False)
    cat['客单价'] = (cat['销售额']/cat['订单数']).round(2)
    # 月度趋势与增长
    m = od.set_index('订单日期').resample('ME')['订单金额'].sum()
    growth = (m.iloc[-1] / m.iloc[0] - 1) * 100
    # 城市 Top
    city = od.groupby('发货城市')['订单金额'].sum().sort_values(ascending=False)
    return {'total': total, 'orders': orders, 'aov': aov,
            'top_cat': cat.head(1).index[0], 'growth': growth,
            'top_city': city.index[0], 'city_top': city.head(3)}

res = analyze_sales(od)
print('===== 电商经营分析结论 =====')
print(f'1. 整体：累计销售额 {res["total"]:,.0f} 元，订单 {res["orders"]:,} 单，客单价 {res["aov"]:,.2f} 元/单')
print(f'2. 品类：销售额最高的是【{res["top_cat"]}】品类')
print(f'3. 趋势：全年销售额从年初到年末增长 {res["growth"]:.1f}%')
print(f'4. 地域：销售额第一的城市是【{res["top_city"]}】，Top3 = {list(res["city_top"].index)}')
""")
code("""# 9.2 生成"一图看懂"总览：2x2 组合图（复用上面函数的结果 + 数据）
fig, axes = plt.subplots(2, 2, figsize=(13, 9))
# 月度趋势
m = od.set_index('订单日期').resample('ME')['订单金额'].sum()
m.index = m.index.strftime('%Y-%m')
axes[0,0].plot(m.index, m.values, marker='o', color='#2E86AB')
axes[0,0].set_title('月度销售趋势'); axes[0,0].tick_params(axis='x', rotation=45)
# 品类占比
cat = od.groupby('品类')['订单金额'].sum().sort_values(ascending=False)
axes[0,1].pie(cat.values, labels=cat.index, autopct='%.1f%%')
axes[0,1].set_title('品类销售占比')
# 城市 Top
city = od.groupby('发货城市')['订单金额'].sum().sort_values(ascending=False).head(6)
axes[1,0].barh(city.index[::-1], city.values[::-1], color='#4C9F70')
axes[1,0].set_title('销售额 Top 城市')
# 客单价分布
axes[1,1].hist(od['订单金额'], bins=40, color='#D98324', edgecolor='white')
axes[1,1].set_title('订单金额分布')
plt.tight_layout(); plt.show()
""")
md("""
### 9.3 大厂面试高频问答（数据分析方向）

**Q1：拿到一份脏数据，你的清洗流程是什么？**
> 答：先"摸结构"（shape/info/describe/unique 看规模、类型、缺失、口径）→ 再逐类处理：缺失值（按业务用均值/中位数/众数填充或删除）、重复值（全字段重复去重、主键重复按规则保留）、异常值（业务规则 + 箱线图 + IQR/3σ 判定后裁剪或剔除）、类型转换（to_datetime/to_numeric）、文本口径统一（strip/lower/映射）。每步都记录处理逻辑，保证可复现。

**Q2：客单价和复购率如何提升，你会怎么分析？**
> 答：先拆分指标。客单价 = 销售额/订单数，可用关联购买（推荐、满减、组合购）提升；复购率看用户分层（RFM），对"重要保持客户"做召回、对"一般发展客户"做促活。用数据定位到具体品类/城市/人群再给策略。

**Q3：groupby 和 SQL 的 group by 区别？**
> 答：逻辑一致——按某字段分组后对各组做聚合（sum/count/mean）。pandas 用 `df.groupby('品类').agg(销售额=('金额','sum'))`，SQL 用 `SELECT 品类, SUM(金额) FROM t GROUP BY 品类`。pandas 更灵活（可对每列用不同聚合），SQL 更标准化、适合大表。

**Q4：怎么判断两个变量的相关性？**
> 答：先画散点图看趋势，再算相关系数（Pearson 适合线性，Spearman 适合单调），最后结合业务解释。相关性≠因果，下结论要谨慎。

**Q5：做 AB 实验要注意什么？**
> 答：样本量要够（可做功效分析）、随机分组要均匀、实验组/对照组只差一个变量、要跑足够时间覆盖周期效应、看显著性和置信区间而非只看均值。
""")
md("""
### 9.4 自测题（检验学习成果）

**第 1 题**：用代码统计"每个月、每个品类的销售额"，并选出销售额最高的月份和品类。
**第 2 题**：把"客户类型"（RFM 结果）为"重要价值客户"的用户筛出来，看他们的城市分布。
**第 3 题**：画一张"各城市客单价"的柱状图，并说明哪个城市客单价最高、可能原因。
**第 4 题**：写出判断"订单金额是否存在缺失"的两种方法（布尔统计 / 可视化）。
**第 5 题**：把订单表按"用户ID"分组，找出消费总金额最高的前 5 位用户。

> 提示：第 1 题可用 `groupby(['订单月份','品类'])`；第 3 题用 `groupby('发货城市')` 算均值再 `barplot`。
""")
code("""# 自测题参考实现（前 3 题示范，先自己思考再看）
# 第1题：每月每品类销售额
q1 = od.groupby(['订单月份', '品类'])['订单金额'].sum().sort_values(ascending=False)
print('每月每品类销售额 Top5：')
print(q1.head(5).round(0))

# 第2题：RFM 重要价值客户的城市分布（需先跑模块5.4得到 rfm 与客户类型）
rfm_with_city = rfm.merge(ui[['用户ID', '城市']], on='用户ID', how='left')
valuable = rfm_with_city[rfm_with_city['客户类型'] == '重要价值客户']
print('\\n重要价值客户城市分布：')
print(valuable['城市'].value_counts().head())

# 第3题：各城市客单价柱状图
city_aov = od.groupby('发货城市')['订单金额'].mean().sort_values(ascending=False)
plt.figure(figsize=(9, 5))
city_aov.plot(kind='bar', color='#4C9F70')
plt.title('各城市客单价对比', fontsize=14)
plt.ylabel('客单价(元)'); plt.xticks(rotation=45)
plt.tight_layout(); plt.show()
""")

# ============================================================
# 模块10：总结
# ============================================================
md("""---
## 模块 10 · 总结与面试考点回顾

### 🎯 你已经掌握的技能点清单
| 类别 | 具体技能 |
|---|---|
| 数据读取 | `read_excel` 多子表、路径管理 |
| 数据概览 | `shape/info/describe/dtypes/nunique/unique` |
| 数据清洗 | 缺失值填充、重复值去重、异常值(IQR/3σ/业务规则)、`to_datetime`、文本清洗 |
| 数据整理 | 列操作、布尔筛选、`query`、`sort_values`、`merge`/`concat`、计算列 |
| 聚合分析 | `groupby`/`agg`/`pivot_table`/`crosstab`/`resample` 环比同比 |
| 业务指标 | 销售额/客单价/复购率/RFM 分层/帕累托 |
| matplotlib | 折线/柱状/饼图/直方图/箱线/散点/子图/样式 |
| seaborn | 分布/热力图/箱线/小提琴/`pairplot`/相关矩阵 |
| plotly | 交互折线/柱状/散点/热力图 |
| 综合分析 | 指标口径、AB 实验、业务洞察、面试问答 |

### 🚀 进阶方向（面试加分）
1. **SQL 版本**：把同样的聚合用 SQL 写一遍（`GROUP BY`、`窗口函数`），对比两者。
2. **机器学习**：用清洗后的数据做"用户是否会再次购买"的分类预测（逻辑回归/XGBoost）。
3. **可视化进阶**：用 plotly 做交互式 BI 看板；学习 Tableau/Power BI 联动。
4. **AB 实验**：学假设检验（t 检验/卡方），做差异显著性分析。

> 💡 **项目复盘建议**：把本项目整理成一份"作品集"（含数据、代码、结论图），放 GitHub，面试时直接展示——这就是大厂面试官想看到的"完整闭环"能力。
""")

nb.cells = cells
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '02_学习任务与代码.ipynb')
with open(OUT, 'w', encoding='utf-8') as f:
    nbf.write(nb, f)
print('Notebook 已生成：', OUT)
print('总单元格数：', len(cells))
