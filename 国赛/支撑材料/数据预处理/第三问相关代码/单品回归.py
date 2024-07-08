import pandas as pd
from datetime import datetime
from sklearn.linear_model import LinearRegression
# TODO:求解6-24到6-30的单品的销售量和销售价格的关系
# 加载Excel数据
data = pd.read_excel("../第二问结果/合并的总数据.xlsx")

# 将'销售日期'列转换为日期格式
data['销售日期'] = pd.to_datetime(data['销售日期'])

# 从原始数据中筛选出在指定日期范围内的数据
selected_data = data[(data['销售日期'] >= '2023-06-24') & (data['销售日期'] <= '2023-06-30')]
# 计算每种商品的平均批发价格
avg_wholesale_prices = selected_data.groupby('单品编码')['批发价格(元/千克)'].mean()
# 获取每种商品的损耗率
product_loss_rates = data.groupby('单品编码')['损耗率(%)'].first()
results = []  # 存放回归结果
# 修改这行代码
for product, avg_price in avg_wholesale_prices.items():
    product_data = data[data['单品编码'] == product]
    X = product_data['销售单价(元/千克)'].values.reshape(-1, 1)  # 特征变量 (价格)
    y = product_data['销量(千克)']                             # 目标变量 (销量)
    model = LinearRegression()
    model.fit(X, y)
    # 计算进货成本
    cost = avg_price / (1 - product_loss_rates[product]/100)
    results.append({
        '单品编码': product,
        'k (coef_)': model.coef_[0],
        'b (intercept_)': model.intercept_,
        'R^2': model.score(X, y),
        '进货成本': cost
    })

# 将回归结果写入Excel
results_df = pd.DataFrame(results)
with pd.ExcelWriter("../第三问结果/regression_results.xlsx") as writer:
    results_df.to_excel(writer, sheet_name='Regression Results', index=False)
