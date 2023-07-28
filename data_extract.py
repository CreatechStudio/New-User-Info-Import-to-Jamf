import pandas as pd
import numpy as np
import re
import datetime
import os

# 定义输入文件名
file_name = input("请将xlsm首先导出为xlsx，并将其放入和本程序同一文件夹，并输入文件名（不含后缀名）:")
input_file_name = f'{file_name}.xlsx'
print(f'正在读取{input_file_name}...')

# 读取Excel文件
df = pd.read_excel(input_file_name)

# 设置pandas选项以避免科学计数法显示
pd.set_option('display.float_format', lambda x: f'{int(x)}' if np.isfinite(x) else '')

# 定义规则校验函数
def validateData(row):
    invalid_rows = {}

    for index, value in row.items():
            if pd.isna(value) or value == '':
                invalid_rows[index] = row.name + 2
    return invalid_rows

# 存储不符合规则的行的行号和问题所在的表头
invalid_rows = {}

# 遍历每一行并进行规则校验
for index, row in df.iterrows():
    row = row.apply(lambda x: x.strip() if isinstance(x, str) else x)
    row_invalid = validateData(row)
    if row_invalid:
        for key, value in row_invalid.items():
            if value - 1 in invalid_rows:
                invalid_rows[value - 1].append(key)
            else:
                invalid_rows[value - 1] = [key]
    df.loc[index] = row

# 提取表头
headers = df.columns

# 创建空字典存储数据
data_dict = {}

# 遍历每一行
for index, row in df.iterrows():
    # 遍历每个表头和对应的单元格值
    for header in headers:
        value = row[header]
        # 跳过空值
        if pd.isna(value):
            continue
        # 将浮点数转换为整数
        if isinstance(value, float):
            value = int(value)
        # 存储到字典中
        if header not in data_dict:
            data_dict[header] = []
        data_dict[header].append(value)

# 存储不符合规则的行的行号
invalid_row_number_only = list(invalid_rows.keys())

# 写入 Invalid Rows 到 error.txt 文件
if invalid_rows:
    # 输出非空的 Invalid Rows
    print("Invalid rows:")
    for row, headers in invalid_rows.items():
        row += 1
        print(f"Row {row}: {', '.join(headers)}")

    # 存储不符合规则的行的行号
    invalid_row_number_only = list(invalid_rows.keys())

    # 写入 Invalid Rows 到 error.txt 文件
    folder_path = "error_log"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"文件夹 {folder_path} 创建成功！")
    else:
        print(f"文件夹 {folder_path} 已经存在。")
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"error_log/error_{timestamp}.txt"
    with open(filename, "w") as file:
        file.write("Invalid rows:\n")
        for row, headers in invalid_rows.items():
            row += 1
            file.write(f"无效行 {row}: {', '.join(headers)}\n")

    print(f"无效数据已写入 {filename}")
    valid_flag = False
else:
    print("恭喜！ 没有发现错误! 开始运行自动化程序！")
    valid_flag = True

# 公布变量
pub_data_dict = data_dict
pub_valid_flag = valid_flag

# 打印全部原始数据
# for header, values in data_dict.items():
#     print(f"{header}: {values}")
# 打印pub_data_dict
# print(pub_data_dict)
# print(pub_invalid_rows)
# 测试输出特定数据
# print(pub_data_dict['Year_Group'][0])
# print(pub_data_dict['1st_Surname'][0])
# print(pub_data_dict['StuEmail'][0])
# print(pub_replaced_data_dict['txtForename'][0])
# print(pub_replaced_data_dict['Year_Group'][0])
# print(pub_replaced_data_dict['txtForm'][0])
# print(pub_replaced_data_dict['intForm'][0])
# print(pub_replaced_data_dict['Day'][0])
# print(pub_replaced_data_dict['txtStuEmail'][0])
# print(pub_data_dict['SchoolID'][0])
# for i in range(len(pub_data_dict['SchoolID'])):
#             row_data = {header: values[i] for header, values in pub_data_dict.items()}
#             replaced_data = {header: values[i] for header, values in pub_replaced_data_dict.items()}
#             print(row_data)
# row_data = {header: values[2] for header, values in pub_data_dict.items()}
# print(row_data)