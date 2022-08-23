import imp
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os
import math

# 行宽、高 于图片像素比例 100*100 == 800*130


# 图片对象
class image_obj(object):
    path = ""
    name = ""
    num = 0
    def __init__(self, path, name, num):
        self.path = path
        self.name = name
        self.num = num

# 26个字母        
chr_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'] 
# 需要转int的配置字段
to_str_list = ["行高", "列宽", "每行多少列", "最多多少行"]

# 读取配置
cfg = {}
with open("./config.txt", "r", encoding="utf8") as f:
    while True:
        line = f.readline()
        if line == "" or line == "\n":
            break
        line.rstrip().lstrip()
        line = line[:-1]
        list = line.split(":")
        val = list[1]
        if list[0] in to_str_list:
            val = int(list[1])
        cfg[list[0]] = val

        


# 读取图片列表
file_list = []
for root, dirs, files in os.walk(cfg["地址"]):
    for file in files:
        file_path = os.path.join(root, file)
        # 提取名字和数量
        end_idx = file.find(".")
        name = file[: end_idx]
        name_list = name.split("_")
        name = name_list[0]
        num = int(name_list[1])
        new_obj = image_obj(file_path, name, num)
        file_list.append(new_obj)
        
# 创建excel
excel = Workbook()

# 创建总览页
# sheet_list=excel.sheetnames
first_sheet_name = '总览'
excel.create_sheet(first_sheet_name,index=0)
first_sheet = excel[first_sheet_name]
# 调整列宽
for item in range(cfg["每行多少列"]):
    first_sheet.column_dimensions[chr_list[item]].width = cfg["列宽"]
# 计算多少行
file_list_num = len(file_list)
row_num = math.ceil(file_list_num / cfg["每行多少列"])
# 调整行高
for item in range(row_num):
    first_sheet.row_dimensions[item + 1].height = cfg["行高"]
# 开始写入
image_idx = 1
for tmp_image_obj in file_list:
    img = Image(tmp_image_obj.path)
    # img.width = cfg["图片宽"]
    # img.height = cfg["图片高"]
    img.width = cfg["列宽"] * 8
    img.height = cfg["行高"] * 1.3
    # 计算列名
    col_idx = image_idx % cfg["每行多少列"]
    if col_idx == 0:
        col_idx = cfg["每行多少列"]
    col_name = chr_list[col_idx-1]
    # 计算行id
    row_idx = image_idx // cfg["每行多少列"] + 1
    # 写入图片
    first_sheet.add_image(img, str(col_name) + str(row_idx))
    image_idx += 1

# 分页分配
sheet_info = []
# 创建分页
for tmp_image_obj in file_list:
    




# 保存excel
excel.save('./out.xlsx')



































