float_list = [3.14, 2.718, 5.0, 7.8, 10.9]

# 使用列表推导式直接转换（截断小数）
int_list = [int(x) for x in float_list]

print(int_list)  # 输出: [3, 2, 5, 7, 10]