import os
import random
import shutil

# 输入文件夹路径
folder_path = r'D:\lg\BaiduSyncdisk\project\person_code\project_self\steels\data\NEU-DET\train'

# 输出文件夹路径
train_path = r'D:\lg\BaiduSyncdisk\project\person_code\project_self\steels\data\NEU-DET\yolo\images\train'
val_path = r'D:\lg\BaiduSyncdisk\project\person_code\project_self\steels\data\NEU-DET\yolo\images\val'
test_path = r'D:\lg\BaiduSyncdisk\project\person_code\project_self\steels\data\NEU-DET\yolo\images\test'

# 创建输出文件夹
os.makedirs(train_path, exist_ok=True)
os.makedirs(val_path, exist_ok=True)
os.makedirs(test_path, exist_ok=True)

# 获取文件列表
file_list = os.listdir(folder_path)
random.shuffle(file_list)

# 计算数据集划分的索引
num_files = len(file_list)
train_size = int(num_files * 0.8)
val_size = int(num_files * 0.1)
test_size = num_files - train_size - val_size

# 遍历文件列表
for i, file in enumerate(file_list):
    if file.endswith('.jpg'):
        # 图片文件
        file_name = os.path.splitext(file)[0]
        txt_file = file_name + '.txt'

        if i < train_size:
            # 放入训练集
            shutil.copy(os.path.join(folder_path, file), os.path.join(train_path, file))
            shutil.copy(os.path.join(folder_path, txt_file), os.path.join(train_path, txt_file))
        elif i < train_size + val_size:
            # 放入验证集
            shutil.copy(os.path.join(folder_path, file), os.path.join(val_path, file))
            shutil.copy(os.path.join(folder_path, txt_file), os.path.join(val_path, txt_file))
        else:
            # 放入测试集
            shutil.copy(os.path.join(folder_path, file), os.path.join(test_path, file))
            shutil.copy(os.path.join(folder_path, txt_file), os.path.join(test_path, txt_file))