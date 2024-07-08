import numpy as np
import pandas as pd

# TODO: 定义加成率规则，根据不同值的销售量进行加成率的求解，再根据价格求解出总销售价格
def get_markup(category, volume):
    if category == "花叶类":
        return 0.5 if volume > 150 else 0.4
    elif category == "花菜类" or category == "水生根茎类":
        return 0.4 if volume > 20 else 0.35
    elif category == "茄类":
        return 0.45 if volume > 20 else 0.5
    elif category == "辣椒类":
        return 0.52 if volume > 90 else 0.48
    elif category == "食用菌类":
        if volume > 60:
            return 0.45
        elif volume < 90:
            return 0.43
    return 0

# 数据
sales_volumes = [
    [183.2035244, 30.4691529, 25.42625257, 30.7029629, 114.8498429, 71.75270772],
    [186.1067636, 27.68025355, 22.13958015, 31.01205567, 111.375011, 67.665067],
    [123.8932995, 16.18327023, 13.23209253, 18.93737901, 80.18253097, 48.3461396],
    [141.3536332, 15.40191168, 14.95426412, 16.33766004, 78.31667274, 52.30741391],
    [136.8572222, 16.62524309, 15.51437566, 16.92528059, 82.29285958, 57.19928721],
    [132.9592006, 18.06215989, 17.56970101, 16.76847274, 87.89205656, 53.30933178],
    [148.9686828, 19.81328433, 20.70167086, 19.90673853, 95.06322967, 60.90279787]
]
purchase_costs = [
    [3.56020918, 8.722404233, 15.14896823, 5.331869051, 5.106392733, 8.22243793],
    [3.553467357, 8.720151353, 15.14269275, 5.316718008, 5.114136527, 7.894035114],
    [3.566480343, 8.718363996, 15.15239562, 5.291037137, 5.12303391, 7.894035114],
    [3.547412449, 8.719435867, 15.1621047, 5.306861256, 5.128069411, 7.894035114],
    [3.536644158, 8.719461819, 15.17182001, 5.307172476, 5.145876117, 7.894035114],
    [3.522388606, 8.719985402, 15.18154154, 5.278613015, 5.146199574, 7.894035114],
    [3.510805047, 8.719253937, 15.1912693, 5.299974123, 5.153450295, 7.894035114]
]

categories = ["花叶类", "花菜类", "水生根茎类", "茄类", "辣椒类", "食用菌类"]
# 计算
total_profit = 0
prices_matrix = []

# 初始化加成率矩阵
markup_matrix = []

for day_sales, day_costs in zip(sales_volumes, purchase_costs):
    day_prices = []
    day_markups = []  # 每天的加成率列表
    for category, volume, cost in zip(categories, day_sales, day_costs):
        markup = get_markup(category, volume)
        day_markups.append(markup)  # 将加成率添加到每天的列表中
        sale_price = (markup * cost) + cost
        day_prices.append(sale_price)
        profit = (sale_price - cost) * volume
        total_profit += profit
    prices_matrix.append(day_prices)
    markup_matrix.append(day_markups)  # 将每天的加成率列表添加到加成率矩阵中

# 输出到Excel
df = pd.DataFrame(prices_matrix, columns=categories)
df.to_excel("../第二问结果/定价矩阵.xlsx", index=False)

# 输出加成率到Excel
df_markup = pd.DataFrame(markup_matrix, columns=categories)
df_markup.to_excel("../第二问结果/加成率矩阵.xlsx", index=False)

print(f"Total Profit over 7 days: {total_profit}")

