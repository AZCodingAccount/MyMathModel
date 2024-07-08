import pandas as pd

# 加载第一张表
table1 = pd.read_excel('../第三问结果/results.xlsx')  # 请替换为第一张表的实际文件路径

# 加载第二张表
table2 = pd.read_excel('../第三问结果/销量和销售单价.xlsx')  # 请替换为第二张表的实际文件路径

# 加载第三张表
table3 = pd.read_excel('../第三问结果/regression_results.xlsx')  # 请替换为第三张表的实际文件路径
# 合并第一张表和第三张表
merged_table = pd.merge(table1, table3[['单品编码', '进货成本']], on='单品编码', how='left')

# 合并第二张表
merged_table = pd.merge(merged_table, table2[['单品编码', '预测销量', '预测销售单价']], on='单品编码', how='left')

# 将相关列的数据类型转换为数值类型（如果有必要）
merged_table['销售价格'] = pd.to_numeric(merged_table['销售价格'], errors='coerce')  # 转换为数值类型，忽略非数值数据

# 计算利润列
merged_table['利润'] = merged_table['销售价格'] - merged_table['进货成本']


# 打印合并后的表格
merged_table.to_excel('../第三问结果/第三问最终结果.xlsx')
