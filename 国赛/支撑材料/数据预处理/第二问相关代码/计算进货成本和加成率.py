import pandas as pd

# TODO:计算进货成本和加成率
# 从先前的Excel表格读取数据
df = pd.read_excel("../第二问结果/品类的销售总量.xlsx", engine='openpyxl', header=[0, 1], index_col=0)  # 更改为您的文件路径

# 获取所有的分类编码
codes = df.columns.get_level_values(0).unique()

# 复制原始列的顺序，并在适当的位置插入新列
cols_order = []
for col in df.columns:
    if col not in cols_order:
        cols_order.append(col)
        if '平均损耗率' in col:
            cols_order.append((col[0], '进货成本'))
            cols_order.append((col[0], '加成率'))

# 用新的列顺序来扩展DataFrame
for col in cols_order:
    if col not in df.columns:
        df[col] = None

for code in codes:
    # 根据每个编码计算进货成本
    df[(code, '进货成本')] = df[(code, '平均进货价格')] / (1 - df[(code, '平均损耗率')] / 100)

    # 计算加成率
    df[(code, '加成率')] = (df[(code, '平均销售单价')] - df[(code, '进货成本')]) / df[(code, '进货成本')]

# 将结果保留三位小数并保存到新的Excel表格中
df = df[cols_order].round(3)
df.to_excel("../第二问结果/第二问最终数据.xlsx", engine='openpyxl')
