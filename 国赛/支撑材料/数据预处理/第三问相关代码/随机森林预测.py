import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error
import numpy as np

# TODO:使用随机森林预测销量和销售单价
# 从Excel文件中读取数据
df = pd.read_excel('../第二问结果/合并的总数据.xlsx')

# 提取日期，并将其转换为日期格式
df['销售日期'] = pd.to_datetime(df['销售日期'])

# 合并相同单品编码和销售日期的销售量和销售单价并分别求和和平均值
grouped_df = df.groupby(['单品编码', '销售日期']).agg({'销量(千克)': 'sum', '销售单价(元/千克)': 'mean'}).reset_index()

results = []

for product_code in grouped_df['单品编码'].unique():
    data = grouped_df[grouped_df['单品编码'] == product_code]

    # 仅考虑2023年的1月、2月、3月、4月、5月、6月的数据
    data = data[data['销售日期'].dt.year == 2023]
    data = data[data['销售日期'].dt.month.isin([1, 2, 3, 4, 5, 6])]

    # 确保你有足够的数据点
    if len(data) < 10:
        continue

    # 准备数据
    data['天数'] = (data['销售日期'] - data['销售日期'].min()).dt.days  # 将日期转换为天数

    X_sales = data[['天数']].values  # 用于销售量预测的特征
    y_sales = data['销量(千克)'].values

    X_price = data[['天数']].values  # 用于销售单价预测的特征
    y_price = data['销售单价(元/千克)'].values

    # 创建随机森林回归模型并进行销售量预测
    model_sales = RandomForestRegressor(n_estimators=100, random_state=42)  # 你可以调整n_estimators等参数
    model_sales.fit(X_sales, y_sales)

    # 进行销售量预测
    test_date = data['天数'].max() + 1  # 假设要预测的日期是最后日期的下一天
    test_date = np.array([[test_date]])  # 将日期转换为NumPy数组

    predicted_sales = model_sales.predict(test_date)[0]

    # 创建随机森林回归模型并进行销售单价预测
    model_price = RandomForestRegressor(n_estimators=100, random_state=42)  # 你可以调整n_estimators等参数
    model_price.fit(X_price, y_price)

    # 进行销售单价预测
    predicted_price = model_price.predict(test_date)[0]

    # 计算销售量误差和销售单价误差
    sales_error = np.sqrt(mean_squared_error(data['销量(千克)'].tail(1), [predicted_sales]))
    price_error = np.sqrt(mean_squared_error(data['销售单价(元/千克)'].tail(1), [predicted_price]))

    results.append({
        '单品编码': product_code,
        '预测销量': predicted_sales,
        '销量预测偏差': sales_error,
        '预测销售单价': predicted_price,
        '销售单价预测偏差': price_error
    })

predicted_df = pd.DataFrame(results)
predicted_df.drop_duplicates(subset=['单品编码'], inplace=True)  # 删除重复的编码
predicted_df.to_excel('../第三问结果/销量和销售单价.xlsx')