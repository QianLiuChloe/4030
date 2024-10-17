import pandas as pd
import re
import numpy as np


def process_size_column(file_path):
    # 读取Excel文件的第一个工作表，并逐行查找包含'size'的列头
    df = pd.read_excel(file_path, header=None)

    # 查找包含'size'的列名和其所在行
    size_column = None
    header_row = None

    # 逐行查找，直到找到包含'size'的列名
    for i, row in df.iterrows():
        for col in row.index:
            if pd.notna(row[col]) and 'size' in str(row[col]).lower():
                size_column = col
                header_row = i
                break
        if header_row is not None:
            break

    if size_column is None:
        print("没有找到包含'size'的列。")
        return

    # 将找到的header_row设为列名，并重新读取数据
    df.columns = df.iloc[header_row]
    df = df[header_row + 1:].reset_index(drop=True)
    size_column_name = df.columns[size_column]

    # 处理size列，提取数据
    size_data = df[size_column_name].astype(str).apply(lambda x: re.split(r'[xX]', x))

    # 找到最大分割的数量，用于创建新列
    max_splits = size_data.apply(len).max()
    size_cols = [f'size{i + 1}' for i in range(max_splits)]

    # 新建size列，并填充数据，未知数据填充为'Unknown'
    for i in range(max_splits):
        df[size_cols[i]] = size_data.apply(lambda x: x[i] if i < len(x) else 'Unknown').str.extract(
            r'(\d+|\bUnknown\b)', expand=False)
        df[size_cols[i]] = df[size_cols[i]].replace(np.nan, 'Unknown')

    # 计算新列的乘积，新建'SIZE calculation'列
    def calculate_product(row):
        product = 1
        for col in size_cols:
            value = row[col]
            product *= int(value) if value.isdigit() else 1
        return product

    df['SIZE calculation'] = df.apply(calculate_product, axis=1)

    # 保存处理后的数据到新的Excel文件
    output_path = file_path.replace('.xlsx', '_size.xlsx')
    df.to_excel(output_path, index=False)
    print(f"Finished {output_path}")


# 调用函数进行处理
# 请将 'input.xlsx' 替换为您的文件路径
process_size_column('[0, 38, 1066, 420]_0.xlsx')
