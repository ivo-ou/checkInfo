# 作业提交检查
# 作业命名格式 姓名 学号
# 使用学号作为对比，姓名仅供参考
# 示例：张三 31788888    则xuehao_startwith = 3 ；xuehao_len = 8

import os, sys
try:
    import time
except ImportError:
    print("正在安装time扩展")
    res = os.system("pip3 install time >nul")
    if res != 0:
        print("time扩展安装失败")
        sys.exit(1)
    import time
try:
    import xlrd
except ImportError:
    print("正在安装xlrd扩展")
    res = os.system("pip3 install xlrd >nul")
    if res != 0:
        print("xlrd扩展安装失败")
        sys.exit(1)
    import xlrd


indexNAME = 0           # 姓名所在表格的列数（0-n）
indexXH = 1           # 学号所在表格的列数（0-n）


def shili():
    print("示例：")
    print("请选择识别模式：1-学号识别	2-名字识别")
    print("1")
    print("你选择的是模式：1\n请输入学号起始字")
    print("3\n请输入学号长度\n8")
    print("请输入作业的路径\n./作业")
    print("请输入花名册的路径\n./花名册.xlsx")
    print("请输入姓名与学号在花名册中的位置，第一列：0 第二列：1 ( 用空格隔开) 表格中相应位置")
    print("0 1")




def init():
    xuehao_startwith = ''  # 学号起始字母
    xuehao_len = ''  # 学号长度
    print("欢迎使用作业比对系统\n")
    mode = input("请选择识别模式：1-学号识别\t2-名字识别\n")
    while mode > '2':
        mode = input ( "请重新选择识别模式：1-学号识别\t2-名字识别\n" )
    print("你选择的是模式："+mode)
    if mode == '1':
        xuehao_startwith = input("请输入学号起始字母\n")
        xuehao_len = input("请输入学号长度\n")

    path = input("请输入作业的路径\n")
    pathHM = input("请输入花名册的路径\n")
    while '.' not in pathHM:
        pathHM = input ( "请输入花名册的路径\t示例：C:user/nobody/Desketop/花名册.xlsx\n" )
    indexNAME, indexXH = input("请输入姓名与学号在花名册中的位置，第一列：0 第二列：1(用空格隔开）\n").split()

    info = {
        'mode': mode,       # 模式
        'path': path,       # 作业路径
        'pathHM': pathHM,   # 花名册路径
        'xuehao_startwith': xuehao_startwith,   # 学号起始字母
        'xuehao_len': xuehao_len,   # 学号长度
        'indexNAME': indexNAME,     # 姓名索引
        'indexXH': indexXH          # 学号索引
    }
    return info



def readfile(mode, path, xuehao_startwith, xuehao_len):  # 读取作业情况，保存为到列表
    data = []  # 新建空列表
    dirs = os.listdir(path)  # 将作业文件名字输出到dirs
    if mode =='1':
        # 文件名称格式化处理
        for file in dirs:
            if not file.startswith('.'):  # 排除隐藏文件
                indexXH = file.find(xuehao_startwith)  # 查找学号起始位置
                xuehao = file[indexXH:int(indexXH + int(xuehao_len))]  # 筛选出学号
                data.append(str(xuehao))  # 将学号整理成列表
        return data  # 输出学号
    else:
        return dirs


def readHM(pathHM, indexXH, indexNAME): # 读取花名册信息
    indexNAME = int(indexNAME)
    indexXH = int(indexXH)
    biaotou = []
    xuehao = []
    name = []
    name_base = []
    dataHM = xlrd.open_workbook(pathHM)
    table = dataHM.sheet_by_index(0)
    # table = data.sheet_by_name('工作表1')           # 通过名字索引
    # 打印data.sheet_names()可发现，返回的值为一个列表，通过对列表索引操作获得工作表1
    for rowNum in range(table.nrows):           # 行数
        rowVale = table.row_values(rowNum)  # 获取每行的行值
        fomat_name = rowVale[indexNAME]
        if rowNum > 0:
            name_base.append ( rowVale[ indexNAME ] )
            if len(str(fomat_name)) == 2:
                fomat_name = '  '.join(rowVale[indexNAME])
            if str(rowVale[indexXH]).find('.'):
                rowVale[indexXH] = str(rowVale[indexXH]).split('.',1)[0]
            xuehao.append(rowVale[indexXH])
            name.append(fomat_name)
        else:       # 输出表头
            biaotou.append('\t'.expandtabs(1) + rowVale[indexNAME] + '\t'.expandtabs(4) + rowVale[indexXH] + '\t'.expandtabs(4) + "完成状况\t")

    # print(xuehao)
    return xuehao, name, biaotou, name_base


def printlist(title, list):
    biaotou = readHM(info['pathHM'], info['indexXH'], info['indexNAME'])
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


def compare_XH(info):       # 学号匹配模式
    data = readfile(info['mode'], info['path'], info['xuehao_startwith'], info['xuehao_len'])
    xuehao, name, biaotou, filename= readHM(info['pathHM'], info['indexXH'], info['indexNAME'])
    react = []
    undo = []
    unread = []

    # compare
    for num in range(len(xuehao)):
        if xuehao[num] in data:
            react.append(str(name[num]) + " " + str(xuehao[num])+"  已完成")
        else:
            react.append(str(name[num]) + " " + str(xuehao[num]) + "  未完成")
            undo.append(str(name[num]) + " " + str(xuehao[num]) + "  未完成")
    # 检查是否有花名册未收集的文件，或者命名不规范的文件
    for num in range(len(data)):
        if data[num] not in xuehao:
            unread.append(data[num])


    printlist("作业完成情况统计", react)
    printlist("未完成情况", undo)
    print("截止至" + str(time.asctime(time.localtime(time.time()))) + "\n应提交作业数量：{}份 已提交作业数量:{}".format(str(len(react)), str(len(react) - len(undo))) + "份")
    if len(unread) != 0:
        print ( "----------------------" )
        print("未识别文件，请手动处理")
        for num in range(len(unread)):
            print(unread[num])

def compare_XM(info):       # 姓名匹配模式
    data = readfile (info[ 'mode'], info[ 'path' ] , info[ 'xuehao_startwith' ] , info[ 'xuehao_len' ] )    # 获取文件名
    xuehao , name, biaotou, name_base= readHM ( info[ 'pathHM' ] , info[ 'indexXH' ] , info[ 'indexNAME' ] )
    react = [ ]
    undo = [ ]
    unread = [ ]
    flag = False
    status = False
    for num in range(len(name_base)):# 花名册循环
            for j in range(len(data)):
                if flag == False :
                    temp = data[j]
                    for i in range(len(temp)):
                        if temp[i] in name_base[num]:
                            if temp[i:i+len(name_base[num])] == name_base[num]:     # 正则匹配
                                del data[j]
                                flag = True
                                status =True
                                break
                else:
                    flag = False
                    break
            if status == True:
                react.append ( str ( name[ num ] ) + " " + str ( xuehao[ num ] ) + "  已完成" )
                status = False
            else:
                react.append ( str ( name[ num ] ) + " " + str ( xuehao[ num ] ) + "  未完成" )
                undo.append ( str ( name[ num ] ) + " " + str ( xuehao[ num ] ) + "  未完成" )


    printlist("作业完成情况统计", react)
    printlist("未完成情况", undo)
    print("截止至" + str(time.asctime(time.localtime(time.time()))) + "\n应提交作业数量：{}份 已提交作业数量:{}".format(str(len(react)), str(len(react) - len(undo))) + "份")
    if len(data) != 0:
        print ( "----------------------" )
        print("未识别文件，请手动处理")
        for num in range(len(data)):
            print(data[num])

if __name__ == '__main__':
    print("\n作业比对脚本" + str(time.asctime(time.localtime(time.time()))))
    shili()
    info = init()
    print(info['mode'])
    if info['mode'] == "1":
        compare_XH(info)
    elif info['mode'] == "2":
        compare_XM(info)
    else:
        print("不知道什么错误，重新来吧")
