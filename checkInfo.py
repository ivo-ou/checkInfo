# 作业提交检查
# 作业命名格式 姓名 学号
# 使用学号作为对比，姓名仅供参考
# 示例：张三 31788888    则xuehao_startwith = 3 ；xuehao_len = 8

import os, sys
import time
import xlrd

# 参数
path = ""  # 作业路径
pathHM = ""     # 花名册路径
xuehao_startwith = ''  # 学号起始字母
xuehao_len = ''  # 学号长度
indexNAME = 0           # 姓名所在表格的列数（0-n）
indexXH = 1           # 学号所在表格的列数（0-n）

def readfile(path, xuehao_startwith, xuehao_len):  # 读取作业情况，保存为到列表
    data = []  # 新建空列表
    dirs = os.listdir(path)  # 将作业文件名字输出到dirs

    # 文件名称格式化处理
    for file in dirs:
        if not file.startswith('.'):  # 排除隐藏文件
            indexXH = file.find(xuehao_startwith)  # 查找学号起始位置
            xuehao = file[indexXH:int(indexXH + int(xuehao_len))]  # 筛选出学号
            name = file[0:3]  # 文件名前三个字符为名字(仅供参考）
            # print(name)
            data.append(str(xuehao))  # 将学号整理成列表

    return data  # 输出学号


def readHM(pathHM, indexXH, indexNAME): # 读取花名册信息
    biaotou = []
    xuehao = []
    name = []
    dataHM = xlrd.open_workbook(pathHM)
    table = dataHM.sheet_by_index(0)
    # table = data.sheet_by_name('工作表1')           # 通过名字索引
    # 打印data.sheet_names()可发现，返回的值为一个列表，通过对列表索引操作获得工作表1
    for rowNum in range(table.nrows):           # 行数
        rowVale = table.row_values(rowNum)  # 获取每行的行值
        fomat_name = rowVale[indexNAME]
        if rowNum > 0:
            if len(str(fomat_name)) == 2:
                fomat_name = '  '.join(rowVale[indexNAME])
            if str(rowVale[indexXH]).find('.'):
                rowVale[indexXH] = str(rowVale[indexXH]).split('.',1)[0]
            xuehao.append(rowVale[indexXH])
            name.append(fomat_name)
        else:       # 输出表头
            biaotou.append('\t'.expandtabs(1) + rowVale[1] + '\t'.expandtabs(4) + rowVale[0] + '\t'.expandtabs(4) + "完成状况\t")

    # print(xuehao)
    return xuehao, name, biaotou


def printlist(title, list):
    biaotou = readHM(pathHM, indexXH, indexNAME)
    #print("\n")
    print("----------------------")
    print(title + "\n")
    if len(list) == 0:
        print("无" + title)
        print("----------------------")
    else:
        print(biaotou[2][0])            # 输出表头，不需要可注释
        for num in range(len(list)):
            print(list[num])
        print("----------------------")


def compare():
    data = readfile(path, xuehao_startwith, xuehao_len)
    xuehao, name, biaotou = readHM(pathHM, indexXH, indexNAME)
    react = []
    undo = []

    # compare
    for num in range(len(xuehao)):
        if xuehao[num] in data:
            react.append(str(name[num]) + " " + str(xuehao[num])+"  已完成")
        else:
            react.append(str(name[num]) + " " + str(xuehao[num]) + "  未完成")
            undo.append(str(name[num]) + " " + str(xuehao[num]) + "  未完成")

    printlist("作业完成情况统计", react)
    printlist("未完成情况", undo)
    print("截止至" + str(time.asctime(time.localtime(time.time()))) + "\n应提交作业数量：{}份 已提交作业数量:{}".format(str(len(react)), str(len(react) - len(undo))) + "份")


if __name__ == '__main__':
    print("\n作业比对脚本" + str(time.asctime(time.localtime(time.time()))))
    compare()
