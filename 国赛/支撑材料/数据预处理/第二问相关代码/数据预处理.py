import pandas as pd

# 读取数据
df4 = pd.read_excel("../第二问结果/第二问预处理数据.xlsx")
df3 = pd.read_excel("../../附件3.xlsx")

# 将df3的日期转为datetime类型，方便后面的运算
df3['日期'] = pd.to_datetime(df3['日期'])
# 对df3按单品编码进行分组
grouped_df3 = df3.groupby('单品编码')
# 初始化一个缓存
cache = {}
def get_wholesale_price(row):
    product_code = row['单品编码']
    sale_date = pd.to_datetime(row['销售日期'])
    # 首先检查缓存
    if product_code in cache:
        cached_date, cached_price = cache[product_code]
        if sale_date == cached_date:  # 如果日期完全匹配，从缓存返回价格
            return cached_price
    if product_code in grouped_df3.groups:
        # 对于每一个单品编码组，查找日期范围内的匹配行
        matched_rows = grouped_df3.get_group(product_code)
        # 过滤出在指定日期范围内的匹配行
        matched_rows = matched_rows[abs(sale_date - matched_rows['日期']) <= pd.Timedelta(days=7)]
        # 按日期的绝对差值升序排序，确保最接近销售日期的行排在最前面
        matched_rows = matched_rows.assign(diff=abs(sale_date - matched_rows['日期'])).sort_values(by='diff')
        if not matched_rows.empty:
            price = matched_rows.iloc[0]['批发价格(元/千克)']
            cache[product_code] = (sale_date, price)  # 更新缓存
            return price
    return None
df4['批发价格(元/千克)'] = df4.apply(get_wholesale_price, axis=1)
df4.to_excel('../第二问结果/合并批发价格的数据.xlsx', index=False)

