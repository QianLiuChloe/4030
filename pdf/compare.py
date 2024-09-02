import os
import shutil

# 定义文件夹路径
jpg_folder = r"C:\Users\chloe\OneDrive\桌面\jpg"  # 第一个文件夹路径
xml_folder = r"C:\Users\chloe\OneDrive\桌面\XML"  # 第二个文件夹路径
output_folder = r"C:\Users\chloe\OneDrive\桌面\new"  # 第三个文件夹路径

# 确保输出文件夹存在
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 获取XML文件名（不含扩展名）
xml_filenames = [os.path.splitext(file)[0] for file in os.listdir(xml_folder) if file.endswith('.xml')]

# 遍历第一个文件夹中的JPG文件
for file in os.listdir(jpg_folder):
    # 检查JPG文件名（不含扩展名）是否在XML文件名列表中
    if os.path.splitext(file)[0] in xml_filenames:
        # 复制文件到第三个文件夹
        shutil.copy(os.path.join(jpg_folder, file), os.path.join(output_folder, file))

print("文件复制完成！")
