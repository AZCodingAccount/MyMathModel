import pandas as pd
# TODO:进行缺失值的删除
# 读取数据
data = pd.read_excel('../../附件2.xlsx')
# 确保两个列都是字符串类型
data['销售日期'] = data['销售日期'].astype(str)
data['扫码销售时间'] = data['扫码销售时间'].astype(str)
# 组合两列并将其转化为datetime类型
data['销售日期_扫码销售时间'] = pd.to_datetime(data['销售日期'] + " " + data['扫码销售时间'], format='mixed')
# 找到所有的退货数据
returns = data[data['销量(千克)'] < 0]
# 创建一个空的DataFrame来存储即将被删除的数据
to_remove = pd.DataFrame()
# 对于每一笔退货，找到与之对应的购买数据
for _, row in returns.iterrows():
    purchase = data[
        (data['销售日期_扫码销售时间'] <= row['销售日期_扫码销售时间']) &
        (data['销售日期_扫码销售时间'] >= row['销售日期_扫码销售时间'] - pd.Timedelta(days=30)) &
        (data['单品编码'] == row['单品编码']) &
        (data['销量(千克)'] == -row['销量(千克)'])
        ].sort_values(by="销售日期_扫码销售时间", ascending=False)

    if not purchase.empty:
        to_remove = pd.concat([to_remove, purchase.iloc[:1]], ignore_index=True)  # 只添加最接近的一条购买记录
# 合并退货数据和相应的购买数据，然后从主数据集中删除它们
# 这个是为了避免退货数据影响正常的分析
to_remove = pd.concat([to_remove, returns], axis=0)
data_cleaned = data.drop(to_remove.index)
# 查找没有匹配的退货记录
indices_to_drop = to_remove.index.intersection(returns.index)
not_matched = returns.drop(indices_to_drop)
# 输出不匹配的退货记录
print(not_matched)
data_cleaned['销量(千克)'] = data_cleaned['销量(千克)'].round(3)
data_cleaned['销售单价(元/千克)'] = data_cleaned['销售单价(元/千克)'].round(2)

# 输出被删除的数据作为日志
to_remove.to_excel('../第一问结果/退货数据的删除日志.xlsx', index=False)


# TODO:进行异常值的过滤
data_cleaned.drop(columns=['销售日期_扫码销售时间'], inplace=True)
# 使用groupby方法根据“单品编码”分类，并为每个分类查找异常值
def find_outliers(group):
    # 计算销量的IQR
    Q1_sale = group['销量(千克)'].quantile(0.25)
    Q3_sale = group['销量(千克)'].quantile(0.75)
    IQR_sale = Q3_sale - Q1_sale

    # 定义异常值的范围，我们可以适当调整为2.5倍的IQR，这会导致较少的异常值
    outlier_condition = (
            (group['销量(千克)'] < (Q1_sale - 2.5 * IQR_sale)) |
            (group['销量(千克)'] > (Q3_sale + 2.5 * IQR_sale))
    )

    group['is_outlier'] = outlier_condition
    return group
# 应用函数并获取带有异常值标记的新DataFrame
data_with_outliers = data_cleaned.groupby('单品编码').apply(find_outliers)
# 把异常值保存为日志
outliers_log = data_with_outliers[data_with_outliers['is_outlier'] == True].drop(columns=['is_outlier'])
outliers_log.to_excel('../第一问结果/第一问异常值日志.xlsx', index=False)
# 从数据集中删除异常值
cleaned_data = data_with_outliers[data_with_outliers['is_outlier'] == False].drop(columns=['is_outlier'])
cleaned_data['销售日期'] = pd.to_datetime(cleaned_data['销售日期'])
# 更改日期格式
cleaned_data['销售日期'] = cleaned_data['销售日期'].dt.date

# 更改单品编码格式
cleaned_data['单品编码'] = cleaned_data['单品编码'].apply(lambda x: '{:.0f}'.format(x))
# 将包含多级索引的 Series 转换为 DataFrame
cleaned_data = cleaned_data.reset_index(drop=True)
# 将清理后的数据保存到新的Excel文件中
cleaned_data.to_excel('../第一问结果/第一问预处理数据.xlsx', index=False)

