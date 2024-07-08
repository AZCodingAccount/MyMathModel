import pandas as pd

# TODO:进行第一问第二小问的数据求解
data = pd.read_excel('../第一问结果/第一问预处理数据.xlsx')
data['销售日期'] = pd.to_datetime(data['销售日期'])
# 按照日、周、月进行分组，计算每个单品的销售总量
daily_sales = data.groupby(['销售日期', '单品编码'])['销量(千克)'].sum().reset_index()
weekly_sales = data.groupby([data['销售日期'].dt.to_period('W'), '单品编码'])['销量(千克)'].sum().reset_index()
monthly_sales = data.groupby([data['销售日期'].dt.to_period('M'), '单品编码'])['销量(千克)'].sum().reset_index()
# 描述性统计
desc_stats = data.groupby('单品编码')['销量(千克)'].describe()
desc_stats = desc_stats.sort_values(by='mean', ascending=False)
# 读取单品名称的数据
products = pd.read_excel('../../附件1.xlsx')


# 将单品名称合并到各个数据集中
def merge_product_name(df):
    return pd.merge(df, products[['单品编码', '单品名称']], on='单品编码', how='left')


daily_sales = merge_product_name(daily_sales)
weekly_sales = merge_product_name(weekly_sales)
monthly_sales = merge_product_name(monthly_sales)
desc_stats = merge_product_name(desc_stats.reset_index()).set_index('单品编码')


# 将单品名称移至第一列
def move_product_name_to_first_column(df):
    cols = list(df.columns)
    cols.insert(0, cols.pop(cols.index('单品名称')))
    return df[cols]


daily_sales = move_product_name_to_first_column(daily_sales)
weekly_sales = move_product_name_to_first_column(weekly_sales)
monthly_sales = move_product_name_to_first_column(monthly_sales)
desc_stats = move_product_name_to_first_column(desc_stats)

# 输出到Excel的不同sheet中，并处理日期格式
with pd.ExcelWriter('../第一问结果/单品汇总结果.xlsx',
                    date_format='YYYY-MM-DD',
                    datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
    daily_sales.to_excel(writer, sheet_name='日销售总量', index=False)
    weekly_sales.to_excel(writer, sheet_name='周销售总量', index=False)
    monthly_sales.to_excel(writer, sheet_name='月销售总量', index=False)
    desc_stats.to_excel(writer, sheet_name='描述性统计', index=True)
