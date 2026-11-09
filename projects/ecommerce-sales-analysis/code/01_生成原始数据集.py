# -*- coding: utf-8 -*-
"""
01_生成原始数据集.py
生成一份"电商销售原始数据"Excel 数据集，共 3 张表：
  1. 订单明细表  order_detail   —— 每一笔订单的销售记录
  2. 用户信息表  user_info      —— 下单用户的基本档案
  3. 商品信息表  product_info   —— 在售商品的基本档案
特意注入"脏数据"（缺失值/重复值/异常值/格式不一致/文本大小写混杂），
用于 Python 数据分析实战中的数据清洗练习。种子固定，可复现。
"""
import numpy as np
import pandas as pd
import random
import os

# ---------- 固定随机种子，保证结果可复现 ----------
rng = np.random.default_rng(2024)
random.seed(2024)

# ============================================================
# 一、构造商品池（约 180 个商品）
# ============================================================
CATEGORIES = ['数码', '家电', '服饰', '食品', '美妆', '图书']
BRANDS = {
    '数码': ['华为', '小米', '苹果', '联想', 'OPPO', 'vivo'],
    '家电': ['美的', '格力', '海尔', 'TCL', '苏泊尔'],
    '服饰': ['优衣库', 'ZARA', '海澜之家', '太平鸟', '李宁', '安踏'],
    '食品': ['三只松鼠', '良品铺子', '百草味', '来伊份'],
    '美妆': ['完美日记', '花西子', '欧莱雅', '雅诗兰黛'],
    '图书': ['中信出版', '人民邮电', '机械工业', '中华书局'],
}
PRODUCT_NAMES = {
    '数码': ['智能手机', '蓝牙耳机', '智能手表', '平板电脑', '笔记本电脑', '无线充电器'],
    '家电': ['空调', '冰箱', '洗衣机', '电饭煲', '空气净化器', '吸尘器'],
    '服饰': ['T恤', '牛仔裤', '羽绒服', '连衣裙', '运动鞋', '衬衫'],
    '食品': ['坚果礼盒', '牛肉干', '薯片', '巧克力', '酸奶', '面包'],
    '美妆': ['口红', '面膜', '精华液', '洗面奶', '防晒霜', '香水'],
    '图书': ['Python编程', '数据分析入门', '统计学原理', '机器学习', '算法导论', '深度学习'],
}

products = []
pid = 1001
for cat in CATEGORIES:
    for name in PRODUCT_NAMES[cat]:
        # 每个品类下的具体款式做 5 种价格档位
        for _ in range(5):
            brand = random.choice(BRANDS[cat])
            if cat == '数码':
                cost = round(random.uniform(300, 5000), 2)
            elif cat == '家电':
                cost = round(random.uniform(200, 4000), 2)
            elif cat == '服饰':
                cost = round(random.uniform(30, 600), 2)
            elif cat == '食品':
                cost = round(random.uniform(10, 150), 2)
            elif cat == '美妆':
                cost = round(random.uniform(40, 500), 2)
            else:  # 图书
                cost = round(random.uniform(15, 120), 2)
            launch = pd.Timestamp('2023-01-01') + pd.Timedelta(
                days=int(rng.integers(0, 700)))
            products.append({
                '商品ID': pid, '商品名称': f'{brand}{name}',
                '品类': cat, '品牌': brand,
                '成本价': cost, '上架日期': launch,
                '库存': int(rng.integers(50, 5000)),
            })
            pid += 1
product_info = pd.DataFrame(products)

# ---- 注入脏数据：商品表 3% 品牌缺失、2% 成本缺失 ----
drop_idx = rng.choice(product_info.index, size=int(len(product_info)*0.03), replace=False)
product_info.loc[drop_idx, '品牌'] = np.nan
drop_idx2 = rng.choice(product_info.index, size=int(len(product_info)*0.02), replace=False)
product_info.loc[drop_idx2, '成本价'] = np.nan

# ============================================================
# 二、构造用户池（约 700 个用户）
# ============================================================
SURNAMES = ['王', '李', '张', '刘', '陈', '杨', '黄', '赵', '周', '吴',
            '徐', '孙', '马', '朱', '胡', '郭', '何', '罗', '高', '林']
GIVEN = ['伟', '芳', '娜', '敏', '静', '磊', '军', '洋', '勇', '艳',
         '杰', '娟', '涛', '明', '超', '秀英', '霞', '平', '刚', '桂英']
CITIES = ['广州', '深圳', '北京', '上海', '杭州', '成都', '武汉', '南京',
          '西安', '重庆', '苏州', '长沙', '郑州', '东莞', '佛山', '天津']

users = []
for i in range(700):
    uid = 1001 + i
    name = random.choice(SURNAMES) + random.choice(GIVEN)
    # 性别注入不一致写法：男/M/male、女/F/female，少数缺失
    g = rng.random()
    if g < 0.03:
        gender = np.nan
    elif g < 0.15:
        gender = random.choice(['M', 'F', 'male', 'female', '男 ', ' 女'])
    else:
        gender = random.choice(['男', '女'])
    # 年龄注入异常值：-5、200、150、15 等
    a = rng.random()
    if a < 0.02:
        age = random.choice([-5, 200, 150, 15, 16])
    elif a < 0.05:
        age = np.nan
    else:
        age = int(rng.integers(18, 65))
    # 注册日期混合格式
    r = rng.random()
    rd = pd.Timestamp('2022-01-01') + pd.Timedelta(days=int(rng.integers(0, 1100)))
    if r < 0.08:
        reg = rd.strftime('%Y/%m/%d')          # 斜杠格式字符串
    elif r < 0.14:
        reg = rd.strftime('%Y%m%d')            # 紧凑数字字符串
    else:
        reg = rd                              # 真正的 datetime
    # 会员等级：注入缺失和带空格
    m = rng.random()
    if m < 0.04:
        level = np.nan
    elif m < 0.10:
        level = random.choice(['普通 会员', ' 黄金会员', '铂金会员 '])
    else:
        level = random.choice(['普通会员', '黄金会员', '铂金会员', '钻石会员'])
    city = random.choice(CITIES)
    users.append({'用户ID': uid, '姓名': name, '性别': gender, '年龄': age,
                  '注册日期': reg, '会员等级': level, '城市': city})
user_info = pd.DataFrame(users)

# ============================================================
# 三、构造订单明细（约 2600 行）
# ============================================================
N_ORDER = 2600
order_ids = [f'OD2024{i:05d}' for i in range(1, N_ORDER+1)]

rows = []
for i in range(N_ORDER):
    pid_ = random.choice(product_info['商品ID'].tolist())
    prod = product_info[product_info['商品ID'] == pid_].iloc[0]
    uid = random.choice(user_info['用户ID'].tolist())
    # 日期混合格式
    d = pd.Timestamp('2024-01-01') + pd.Timedelta(days=int(rng.integers(0, 365)))
    r = rng.random()
    if r < 0.06:
        date = d.strftime('%Y/%m/%d')
    elif r < 0.10:
        date = d.strftime('%Y%m%d')
    else:
        date = d
    # 数量：注入负数/0/缺失
    q = rng.random()
    if q < 0.02:
        qty = random.choice([-2, -1, -5])
    elif q < 0.04:
        qty = 0
    elif q < 0.07:
        qty = np.nan
    else:
        qty = int(rng.integers(1, 6))
    # 单价：成本上浮 20%~180% 作为零售价，1% 极端异常值
    if rng.random() < 0.01:
        price = round(random.uniform(10000, 99999), 2)
    else:
        price = round(prod['成本价'] if not pd.isna(prod['成本价']) else 100
                      * random.uniform(1.2, 2.8), 2)
    # 折扣：多数 0~0.3，少数 0.5/0.8 异常
    if rng.random() < 0.03:
        disc = random.choice([0.5, 0.8, 0.6])
    else:
        disc = round(random.uniform(0, 0.3), 2)
    # 运费
    ship = round(random.uniform(0, 20), 2)
    # 订单金额：多数按公式计算，2% 缺失
    if rng.random() < 0.02:
        amount = np.nan
    else:
        amount = round(qty*price*(1-disc) + ship, 2) if not pd.isna(qty) else np.nan
    # 城市：注入大小写/空格/后缀
    c = rng.random()
    if c < 0.05:
        city = random.choice([' 广州', 'Guangzhou', '广州 ', 'shenzhen', '深圳省'])
    else:
        city = random.choice(CITIES)
    # 状态：注入大小写/空格
    s = rng.random()
    if s < 0.06:
        status = random.choice(['已完成 ', ' 已发货', '退款中', 'REFUND', '完成'])
    else:
        status = random.choice(['已完成', '已发货', '待发货', '退款'])
    rows.append({'订单号': order_ids[i], '用户ID': uid, '订单日期': date,
                 '商品ID': pid_, '商品名称': prod['商品名称'], '品类': prod['品类'],
                 '数量': qty, '单价': price, '折扣': disc, '运费': ship,
                 '订单金额': amount, '发货城市': city, '订单状态': status})
order_detail = pd.DataFrame(rows)

# ---- 注入重复订单号：随机挑 25 行整体复制一次，形成完全重复行 ----
dup_rows = order_detail.sample(n=25, random_state=2024)
order_detail = pd.concat([order_detail, dup_rows], ignore_index=True)

# ---- 再注入"订单号重复但数据不同"的假重复（同一订单号出现两次不同内容）----
fake_dup = order_detail.sample(n=15, random_state=7).copy()
fake_dup['订单金额'] = fake_dup['订单金额'] * 2   # 改成不同数据
fake_dup['数量'] = fake_dup['数量'] + 3
order_detail = pd.concat([order_detail, fake_dup], ignore_index=True)

# ---- 打乱行序 ----
order_detail = order_detail.sample(frac=1, random_state=2024).reset_index(drop=True)

# ============================================================
# 四、写出 Excel 文件（3 个子表）
# ============================================================
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '数据集')
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PATH = os.path.join(OUT_DIR, '电商销售原始数据.xlsx')

with pd.ExcelWriter(OUT_PATH, engine='openpyxl') as writer:
    order_detail.to_excel(writer, sheet_name='订单明细表', index=False)
    user_info.to_excel(writer, sheet_name='用户信息表', index=False)
    product_info.to_excel(writer, sheet_name='商品信息表', index=False)

print('文件已生成：', OUT_PATH)
print('订单明细表行数：', len(order_detail))
print('用户信息表行数：', len(user_info))
print('商品信息表行数：', len(product_info))
