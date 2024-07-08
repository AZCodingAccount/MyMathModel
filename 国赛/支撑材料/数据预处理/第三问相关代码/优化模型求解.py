import numpy as np
import pandas as pd
from scipy.optimize import minimize

# TODO:非线性规划，求解出来估计的成本加成定价的关系
# 读取数据
data = pd.read_excel('../第三问结果/regression_results.xlsx')
product_codes = data['单品编码'].tolist()
k = data.set_index('单品编码')['k (coef_)'].to_dict()
b = data.set_index('单品编码')['b (intercept_)'].to_dict()
q = data.set_index('单品编码')['进货成本'].to_dict()


# 定义目标函数和约束
def objective(variables):
    x_values = variables[:len(product_codes)]
    p_values = variables[len(product_codes):2 * len(product_codes)]
    w_values = variables[2 * len(product_codes):3 * len(product_codes)]
    c_values = variables[3 * len(product_codes):]

    profit_values = np.array(w_values) * np.array(x_values)
    total_profit_values = np.array(c_values) * profit_values
    return -np.sum(total_profit_values)


def constraint1(variables):
    return 27 - np.sum(variables[3 * len(product_codes):])


def constraint2(variables):
    return np.sum(variables[3 * len(product_codes):]) - 33


def x_constraints(variables):
    return np.array(variables[:len(product_codes)]) - (np.array([k[code] for code in product_codes]) * np.array(
        variables[len(product_codes):2 * len(product_codes)]) + np.array([b[code] for code in product_codes]))


def p_constraints(variables):
    return np.array(variables[len(product_codes):2 * len(product_codes)]) - (
            1 + np.array(variables[2 * len(product_codes):3 * len(product_codes)])) * np.array(
        [q[code] for code in product_codes])


def w_constraints(variables):
    w_values = variables[2 * len(product_codes):3 * len(product_codes)]
    return 100 - np.max(w_values)


constraints = [
    {'type': 'ineq', 'fun': constraint1},
    {'type': 'ineq', 'fun': constraint2},
    {'type': 'eq', 'fun': x_constraints},
    {'type': 'eq', 'fun': p_constraints},
    {'type': 'ineq', 'fun': w_constraints}  # 新的w_values约束
]

# 设置初始值和边界
initial_guess = [3] * 3 * len(product_codes) + [0] * len(product_codes)
bounds = [(2.5, 500)] * len(product_codes) + [(0, 300)] * len(product_codes) + [(0, 5000)] * len(product_codes) + [
    (0, 1)] * len(product_codes)

# 使用minimize解决问题
result = minimize(objective, initial_guess, bounds=bounds, constraints=constraints, method='SLSQP')

# 创建结果DataFrame
result_data = []

# 输出结果
print("Maximized W总:", -result.fun)
variables = result.x
selected_items_count = 0
for i, code in enumerate(product_codes):
    if variables[3 * len(product_codes) + i] == 1:
        selected_items_count += 1
        result_data.append([code,
                            variables[i],
                            variables[2 * len(product_codes) + i],
                            variables[len(product_codes) + i]])

print(f"\nTotal selected items: {selected_items_count}")
# 创建结果DataFrame
result_df = pd.DataFrame(result_data, columns=['单品编码', '补货量', '利润', '销售价格'])
# 读取单品名称表
product_names = pd.read_excel('../../附件1.xlsx')  # 假设单品名称表的文件名是单品名称表.xlsx
# 将单品名称合并到优化模型结果表中，根据单品编码匹配
merged_results = result_df.merge(product_names[['单品编码', '单品名称']], on='单品编码', how='left')
# 重新排列列的顺序，将单品名称列追加到最前面
merged_results = merged_results[['单品名称'] + [col for col in merged_results.columns if col != '单品名称']]
# 将结果保存到Excel文件
merged_results.to_excel('../第三问结果/results.xlsx', index=False)
