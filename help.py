import os


# 需要转int的配置字段
to_str_list = ["行高", "列宽", "每行多少列", "最多多少行"]


# 获取大写的26个字母
def get_26_char():
    list = []
    for item in range(97, 97+26):
        list.append(str.upper(chr(item)))
    return list

# 获取配置
def get_cfg():
    cfg = {}
    with open("./config.txt", "r", encoding="utf8") as f:
        while True:
            line = f.readline()
            if line == "" or line == "\n":
                break
            line.rstrip().lstrip()  # 去掉多余空格
            line = line[:-1]        # 去掉换行符
            list = line.split(":")
            val = list[1]
            if list[0] in to_str_list:
                val = int(list[1])
            cfg[list[0]] = val
    return cfg














