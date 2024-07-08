import pandas as pd

# TODO:根据日期对不同品类进行，每周，每天，每月的一个销售总量统计
# 读取销售数据
df = pd.read_excel('../第二问结果/剔除异常值后的总数据.xlsx', engine='openpyxl')
# 日期处理
df['销售日期'] = pd.to_datetime(df['销售日期'])
df['日期'] = df['销售日期'].dt.strftime('%Y-%m-%d')
df['周'] = df['销售日期'].dt.to_period('W').astype(str)
df['季度'] = df['销售日期'].dt.to_period('Q').astype(str)
with pd.ExcelWriter('../第二问结果/品类的销售总量.xlsx', engine='openpyxl') as writer:
    for period, label in [('日期', 'Day'), ('周', 'Week'), ('季度', 'Quarter')]:
        # 销售量
        sales_grouped = df.groupby([period, '分类编码', '分类名称'])['销量(千克)'].sum().unstack(level=[1, 2])

        # 平均销售单价
        avg_sales_price_grouped = df.groupby([period, '分类编码'])['销售单价(元/千克)'].mean().unstack(level=1)
        avg_sales_price_grouped.columns = pd.MultiIndex.from_product(
            [avg_sales_price_grouped.columns, ['平均销售单价']],
            names=['分类编码', '属性'])

        # 平均进货价格
        price_grouped = df.groupby([period, '分类编码'])['批发价格(元/千克)'].mean().unstack(level=1)
        price_grouped.columns = pd.MultiIndex.from_product([price_grouped.columns, ['平均进货价格']],
                                                           names=['分类编码', '属性'])

        # 平均损耗率
        wastage_grouped = df.groupby([period, '分类编码'])['损耗率(%)'].mean().unstack(level=1)
        wastage_grouped.columns = pd.MultiIndex.from_product([wastage_grouped.columns, ['平均损耗率']],
                                                             names=['分类编码', '属性'])

        # 重新排序列以使它们的顺序一致
        cols = sales_grouped.columns.tolist()
        reordered_cols = sum(
            [(col, (col[0], '平均销售单价'), (col[0], '平均进货价格'), (col[0], '平均损耗率')) for col in cols], ())
        merged = pd.concat([sales_grouped, avg_sales_price_grouped, price_grouped, wastage_grouped], axis=1).reindex(
            reordered_cols, axis=1)

        # 保存到Excel
        merged.to_excel(writer, sheet_name=label)
