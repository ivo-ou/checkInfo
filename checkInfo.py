# 作业提交检查
# 作业命名格式 姓名 学号
# 使用学号作为对比，姓名仅供参考
# 示例：张三 31788888    则xuehao_startwith = 3 ；xuehao_len = 8

import os, sys
import time

# 参数
path = ""  # 作业路径
xuehao_startwith = ''  # 学号起始字母
xuehao_len = ''  # 学号长度


def readfile(path, xuehao_startwith, xuehao_len):  # 读取作业情况，保存为到列表
    data = []  # 新建空列表
    dirs = os.listdir(path)  # 将作业文件名字输出到dirs

    # 文件名称格式化处理
    for file in dirs:
        if not file.startswith('.'):  # 排除隐藏文件
            indexXH = file.find(xuehao_startwith)  # 查找学号起始位置
            xuehao = file[indexXH:int(indexXH + int(xuehao_len))]  # 筛选出学号
            name = file[0:3]  # 文件名前三个字符为名字(仅供参考）
            print(name)
            data.append(xuehao)  # 将学号整理成列表

    return data  # 输出学号


if __name__ == '__main__':
    print("作业提交比对脚本")
    data = readfile(path, xuehao_startwith, xuehao_len)
    print("\n截止至" + str(time.asctime(time.localtime(time.time()))) + " 已提交作业数量:" + str(
        len(data)) + "人")  # 打印作业提交数量
