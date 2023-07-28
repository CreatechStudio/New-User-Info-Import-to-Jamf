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

# 定义替换规则
replace_dict_year_group = {
    'YEAR 9 (IG) - (1)': 2,
    'YEAR 10(A) - (5)': 6,
    'YEAR 11(A) - (3)': 4,
    'YEAR 10 (IB) - (6)': 7,
    'YEAR 9(Spring) - (9)': 10,
    'TEST - (48)': 14
}

replace_dict_year_group_txt = {
    'YEAR 9 (IG) - (1)': 'G1-',
    'YEAR 10(A) - (5)': 'Pre-',
    'YEAR 11(A) - (3)': 'AS-',
    'YEAR 10 (IB) - (6)': 'Pre-IB',
    'YEAR 9(Spring) - (9)': 'Spring-',
    'TEST - (48)': 'Test-'
}

replace_dict_day ={
    'Boarding': '2',
    'Day': '3'
}

replace_dict_form_index = {
    'G1-1': 2,
    'G1-10': 3,
    'G1-2': 4,
    'G1-3': 5,
    'G1-4': 6,
    'G1-5': 7,
    'G1-6': 8,
    'G1-7': 9,
    'G1-8': 10,
    'G1-9': 11,
    'Pre-1': 2,
    'Pre-10': 3,
    'Pre-2': 4,
    'Pre-3': 5,
    'Pre-4': 6,
    'Pre-5': 7,
    'Pre-6': 8,
    'Pre-7': 9,
    'Pre-8': 10,
    'Pre-9': 11,
    'AS-1': 2,
    'AS-10': 3,
    'AS-11': 4,
    'AS-12': 5,
    'AS-13': 6,
    'AS-2': 7,
    'AS-3': 8,
    'AS-4': 9,
    'AS-5': 10,
    'AS-6': 11,
    'AS-7': 12,
    'AS-8': 13,
    'AS-9': 14,
    'Pre-IB1': 3,
    'Pre-IB2': 4,
    'Spring-1': 2,
    'Spring-2': 3,
    'Spring-3': 4,
    'Spring-4': 5,
    'Spring-5': 6,
    'Spring-6': 7,
    'Test-1': 2
}

# 创建新的字典存储替换后的值
replaced_data_dict = {}

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
pub_replaced_data_dict = replaced_data_dict
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