import pandas as pd
import numpy as np

# 从Excel文件中读取数据
df = pd.read_excel('../第一问结果/单品汇总结果.xlsx', sheet_name='周销售总量', engine='openpyxl')
# 创建单品编码到单品名称的映射
code_to_name = dict(zip(df['单品编码'], df['单品名称']))

# 使用pivot方法重塑数据
pivot_df = df.pivot(index='销售日期', columns='单品编码', values='销量(千克)')
# 查看并打印填充0之前有多少个空值
missing_values = pivot_df.isna().sum().sum()  # 这里两次使用sum()：第一次是计算每列的NaN数量，第二次是计算所有列的总数。

# 使用0填充pivot_df中的缺失值
pivot_df.fillna(0, inplace=True)

# 计算相关性
correlation = pivot_df.corr(method='spearman')

# 更改相关性矩阵的列名和行标签
correlation.rename(columns=code_to_name, index=code_to_name, inplace=True)

# 使用上三角矩阵筛选相关系数大于0.8或小于-0.8的值
upper_tri = correlation.where(np.triu(np.ones(correlation.shape), k=1).astype(bool))

# 为了避免列名冲突，我们在stack之前重命名列级别
upper_tri.columns.name = "Product Code"
strong_corr_pairs = upper_tri.stack().reset_index()
strong_corr_pairs.columns = ['单品1', '单品2', '相关系数']
strong_corr_pairs = strong_corr_pairs[(strong_corr_pairs['相关系数'] > 0.8) | (strong_corr_pairs['相关系数'] < -0.8)]

# 根据相关系数的绝对值对强相关对进行排序
strong_corr_pairs['abs_correlation'] = strong_corr_pairs['相关系数'].abs()
strong_corr_pairs = strong_corr_pairs.sort_values(by='abs_correlation', ascending=False).drop(columns=['abs_correlation'])
# 输出到Excel文件
with pd.ExcelWriter("../第一问结果/第二小问相关系数结果.xlsx", engine='openpyxl') as writer:
    correlation.to_excel(writer, sheet_name="相关系数矩阵")
    strong_corr_pairs.to_excel(writer, sheet_name="强相关系数")
