from openpyxl.drawing.image import Image
import os

import help as Help


# 26个字母  
chr_list = Help.get_26_char()
# 读取配置
cfg = Help.get_cfg()

# 图片对象
class ImageObj(object):
    path = ""
    name = ""
    num = 0
    def __init__(self, path, name, num):
        self.path = path
        self.name = name
        self.num = num

# 获取目录下的文件图片对象
def get_dir_image_list():
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
            new_obj = ImageObj(file_path, name, num)
            file_list.append(new_obj)
    return file_list

# 创建sheet
def create_sheet(excel, sheet_name, sheet_index):
    excel.create_sheet(sheet_name,index=sheet_index)
    sheet = excel[sheet_name]
    # 调整列宽
    for item in range(cfg["每行多少列"]):
        sheet.column_dimensions[chr_list[item]].width = cfg["列宽"]
    # 调整行高
    for item in range(cfg["最多多少行"]):
        sheet.row_dimensions[item + 1].height = cfg["行高"]
    return sheet

# 为sheet插入图片
def sheet_add_image(sheet, image_idx, path):
    img = Image(path)
    img.width = cfg["列宽"] * 8
    img.height = cfg["行高"] * 1.3
    col_name,row_idx = get_col_row_by_idx(image_idx)
    sheet.add_image(img, col_name + str(row_idx))

# 通过idx获取列和行编号
def get_col_row_by_idx(idx):
    # 计算行id
    row_idx = idx // cfg["每行多少列"] + 1
    # 计算列名
    col_idx = idx % cfg["每行多少列"]
    if col_idx == 0:
        col_idx = cfg["每行多少列"]
        row_idx -= 1 
    col_name = chr_list[col_idx-1]
    
    return col_name,row_idx
 
    

























