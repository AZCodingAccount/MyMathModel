import pandas as pd
import numpy as np
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
import matplotlib.pyplot as plt

# 从Excel读取数据
file_path = 'D:\myPython\myProject\数据预处理\第一问结果\单品汇总结果.xlsx'  # 替换为你的Excel文件路径
df = pd.read_excel(file_path, engine='openpyxl')

# 与之前类似地处理数据
df_grouped = df.groupby(['单品编码', '销售日期'])['销量(千克)'].sum().reset_index()
df_grouped['销售日期'] = pd.to_datetime(df_grouped['销售日期'])
ref_date = min(df_grouped['销售日期'])
df_grouped['天数'] = (df_grouped['销售日期'] - ref_date).dt.days

# 准备用于聚类的数据
X = df_grouped[['天数', '销量(千克)']].values

# 层次聚类
linked = linkage(X, 'ward')

# 绘制树状图
plt.figure(figsize=(10, 7))
dendrogram(linked, orientation='top', distance_sort='descending', show_leaf_counts=True)
plt.show()

# 根据树状图切割聚类
max_d = 50000  # 通过增加这个值来减少聚类的数量
clusters = fcluster(linked, max_d, criterion='distance')
df_grouped['Cluster'] = clusters

# 对每个聚类提取和去重单品名称
clustered_items = []
for cluster in np.unique(clusters):
    unique_items = df[df_grouped['Cluster'] == cluster]['单品名称'].unique()
    clustered_items.append([f"Cluster_{cluster}"] + unique_items.tolist())

# 将数据转换为一个DataFrame并保存到Excel
df_output = pd.DataFrame(clustered_items).fillna('')
df_output.to_excel('clusters_hierarchical_output.xlsx', header=False, index=False)

print("Data saved to clusters_hierarchical_output.xlsx")
