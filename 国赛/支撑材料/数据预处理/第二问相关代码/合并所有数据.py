import pandas as pd

# TODO:损耗率和分类名称均是根据单品编码合并，因此就不再赘述，使用excel命令VLookup函数合并
# 1. 读取原始数据
data = pd.read_excel("../第二问结果/合并的总数据.xlsx")

# 2. 使用条件筛选数据
# 找到销售单价大于批发价格50倍的数据
filtered_data_indices = data[data['销售单价(元/千克)'] > 20 * data['批发价格(元/千克)']].index

# 3. 从原始数据中删除滤出的数据
data.drop(filtered_data_indices, inplace=True)

# 4. 保存处理后的数据到新的Excel文件
data.to_excel("../第二问结果/剔除异常值后的总数据.xlsx", index=False)
