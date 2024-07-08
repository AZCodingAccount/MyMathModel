import pandas as pd

# 读取分类数据
df1 = pd.read_excel('../../附件1.xlsx', engine='openpyxl')

# 提取分类编码及其对应的单品编码和分类名称
result = df1.groupby('分类编码').agg({'单品编码': list, '分类名称': 'first'}).reset_index()

# 创建单品编码到分类编码和分类名称的映射字典
class_code_dict = dict()
class_name_dict = dict()

for index, row in result.iterrows():
    for item_code in row['单品编码']:
        class_code_dict[item_code] = row['分类编码']
        class_name_dict[item_code] = row['分类名称']

# 读取销售数据
df = pd.read_excel('../第一问结果/第一问预处理数据.xlsx', engine='openpyxl')

# 为数据添加分类编码和分类名称列
df['分类编码'] = df['单品编码'].map(class_code_dict)
df['分类名称'] = df['单品编码'].map(class_name_dict)

# 日期处理
df['销售日期'] = pd.to_datetime(df['销售日期'])

# 按天分组
df['日期'] = df['销售日期'].dt.strftime('%Y-%m-%d')
grouped_day = df.groupby(['日期', '分类编码', '分类名称'])['销量(千克)'].sum().unstack(level=[1, 2])

# 按周分组
df['周'] = df['销售日期'].dt.strftime('%Y-W%U')
grouped_week = df.groupby(['周', '分类编码', '分类名称'])['销量(千克)'].sum().unstack(level=[1, 2])

# 按季度分组
df['季度'] = df['销售日期'].dt.year.astype(str) + '-Q' + ((df['销售日期'].dt.month - 1) // 3 + 1).astype(str)
grouped_quarter = df.groupby(['季度', '分类编码', '分类名称'])['销量(千克)'].sum().unstack(level=[1, 2])

# 保存到一个Excel文件中的三个不同sheets
with pd.ExcelWriter('../第一问结果/分类汇总结果.xlsx') as writer:
    grouped_day.to_excel(writer, sheet_name='按天分类')
    grouped_week.to_excel(writer, sheet_name='按周分类')
    grouped_quarter.to_excel(writer, sheet_name='按季度分类')
