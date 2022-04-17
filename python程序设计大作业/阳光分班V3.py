"""
 -*- coding: utf-8 -*-

 @Time : 2021/11/10 18:02

 @Author : jagger

 @File : 阳光分班.py

 @Software: PyCharm 

 @contact: 252587809@qq.com

 -*- 功能说明 -*-

"""
import random
import webbrowser
import pandas as pd
from tkinter import *
from tkinter import messagebox
from PIL import Image, ImageTk  # PIL图像处理库

STUDENT = 1200  # 学生总人数
CLASS_COUNT = 22  # 班级总数
name_dict = []  # 该年级学生信息总列表
sex = [0, 0]  # 学生男女性别数量,左男右女
CLASS_SEX = [[0, 0]]  # 各班男女数量,下标0无用


# 打开Excel表格
def open_word():
    webbrowser.open('C:\Python\src\python作业\阳光分班2.xlsx')


# 提示是否要关闭窗口
def check(windows):
    if (messagebox.askokcancel('提示', '要执行此操作吗') == 1):
        windows.destroy()  # 关闭窗口


# 设置图像合适大小
def get_image(filename, width, height):
    im = Image.open(filename).resize((width, height))  # resize调整图片大小
    return ImageTk.PhotoImage(im)


'''
def Unicode():
    val = random.randint(0x4e00, 0x9fbf)
    return chr(val)
'''


# 随机选取名,gbk2312对字符的编码采用两个字节相组合,第一个字节的范围是0xB0-0xF7, 第二个字节的范围是0xA1-0xFE.
def GBK2312():
    head = random.randint(0xb0, 0xf7)
    body = random.randint(0xa1, 0xf9)
    val = f'{head:x}{body:x}'
    str = bytes.fromhex(val).decode('gb2312')
    return str


# 创建 1200 个学生姓名,性别,年龄,语文,数学,英语,总分,和初始班级号0
def create_student():
    # 百家姓
    str1 = ["王", "李", "张", "刘", "陈", "杨", "黄", "吴", "赵", "周", "徐", "孙", "马", "朱", "胡", "林", "郭", "何", "高", "罗", "郑",
            "梁", "谢", "宋",
            "唐", "许", "邓", "冯", "韩", "曹", "曾", "彭", "萧", "蔡", "潘", "田", "董", "袁", "于", "余", "叶", "蒋", "杜", "苏", "魏",
            "程", "吕", "丁",
            "沈", "任", "姚", "卢", "傅", "钟", "姜", "崔", "谭", "廖", "范", "汪", "陆", "金", "石", "戴", "贾", "韦", "夏", "邱", "方",
            "侯", "邹", "熊",
            "孟", "秦", "白", "江", "阎", "薛", "尹", "段", "雷", "黎", "史", "龙", "陶", "贺", "顾", "毛", "郝", "龚", "邵", "万", "钱",
            "严", "赖", "覃",
            "洪", "武", "莫", "孔", "鐘", "司马", "上官", "欧阳", "夏侯", "诸葛", "东方", "皇甫", "公孙", "轩辕", "令狐", "司徒", "宇文"]

    # 创建学生信息
    for i in range(STUDENT):
        # 随机选取姓
        name = ''
        name += str1[random.randint(0, len(str1) - 1)]

        # 随机选取名
        for j in range(random.randint(1, 2)):
            name += GBK2312()

        # 随机选取性别
        s = random.choice(('男', '女'))
        if s == '男':
            sex[0] += 1
        else:
            sex[1] += 1

        # 随机选取年龄
        age = random.randint(11, 12)

        # 随机选取语文成绩
        Chinese = random.randint(0, 100)

        # 随机选取数学成绩
        Math = random.randint(0, 100)

        # 随机选取英语成绩
        English = random.randint(0, 100)

        # 三科总分
        total = Chinese + Math + English

        # 初始班级号为0
        class_son = 0

        # 平均分
        # 两数相除，结果只保留一位小数点
        # avg = round(total/3,1)

        # 将学生信息以列表形式放入列表中,整体为列表嵌套
        name_dict.append([name, s, age, Chinese, Math, English, total, class_son])


def Fen_Pei():
    # 分配每个班的男女比例
    boy = sex[0] // CLASS_COUNT
    girl = sex[1] // CLASS_COUNT
    var1 = []
    var2 = []
    for i in range(1, CLASS_COUNT + 1):
        CLASS_SEX.append([boy, girl])
        var1.append(i)
        var2.append(i)
    for i in range(sex[0] % CLASS_COUNT):
        h = random.choice(var1)
        CLASS_SEX[h][0] += 1
        var1.remove(h)
    for i in range(sex[1] % CLASS_COUNT):
        h = random.choice(var2)
        CLASS_SEX[h][1] += 1
        var2.remove(h)


# 排序函数
def Sort():
    # 根据分配好的班级号排序,从1~22班,升序排序
    name_dict.sort(key=lambda item: item[7])


# 分班函数
def divison_class():
    var = []
    for i in range(1, CLASS_COUNT + 1):
        var.append(i)

    for i in range(STUDENT):
        if (name_dict[i][1] == '男'):
            while True:
                h = random.choice(var)
                if (CLASS_SEX[h][0] != 0):
                    break
            name_dict[i][7] = h
            CLASS_SEX[h][0] -= 1
        else:
            while True:
                h = random.choice(var)
                if (CLASS_SEX[h][1] != 0):
                    break
            name_dict[i][7] = h
            CLASS_SEX[h][1] -= 1


# 将分配好班级的学生导入Excel表格
def write_excel():
    begin = 0
    end = 0
    stop = 0
    class_son = 1
    writer = pd.ExcelWriter(r'C:\Python\src\python作业\阳光分班2.xlsx')
    for i in range(STUDENT):
        if (name_dict[i][7] == class_son):
            end += 1
            stop += 1
        else:
            data = pd.DataFrame(name_dict[begin:end:1])
            data.columns = [u'名字', u'性别', u'年龄', u'语文', u'数学', u'英语', u'总分', u'班级号']
            data.index = pd.RangeIndex(start=1, stop=stop + 1, step=1)
            class_name = str('%d班' % class_son)
            data.to_excel(writer, sheet_name=class_name)
            begin = end
            end += 1
            stop = 1
            class_son += 1

    data = pd.DataFrame(name_dict[begin:end:1])
    data.columns = [u'名字', u'性别', u'年龄', u'语文', u'数学', u'英语', u'总分', u'班级号']
    data.index = pd.RangeIndex(start=1, stop=stop + 1, step=1)
    class_name = str('%d班' % class_son)
    data.to_excel(writer, sheet_name=class_name)

    writer.save()


def body():
    create_student()
    Fen_Pei()
    divison_class()
    Sort()
    write_excel()
    messagebox.showinfo('提示', '完成分班')


if __name__ == '__main__':
    # 设置窗口，静止放大缩小，设置标题，设置窗体大小和在屏幕的位置
    windows = Tk()
    windows.title("阳光分班程序")
    windows.resizable(False, False)
    windows.geometry('400x200+500+200')

    # 创建画布，将图片放在画布上
    canvas = Canvas(windows, width=400, height=200)
    im = get_image('logo.png', 400, 200)
    canvas.create_image(200, 100, image=im)
    canvas.create_text(210, 55, fill='green', text='阳光分班程序', font=('华文行楷', 30))
    canvas.pack()

    buttom1 = Button(canvas, text='开始分班', bg='lightgreen', fg='black', font=('黑体', 15),
                     command=lambda: body())  # lambda（）函数用于将特定数据发送到回调函数。
    buttom1.lift  # 将按钮上调到主界面不被Canvas覆盖
    buttom1.place(x=60, y=130)

    buttom2 = Button(canvas, text='查看结果', bg='lightgreen', fg='black', font=('黑体', 15), command=lambda: open_word())
    buttom2.lift
    buttom2.place(x=260, y=130)

    # protocol协议，WM_DELETE_WINDOW窗体关闭,这段代码是让窗体右上角x按钮点击关闭
    windows.protocol('WM_DELETE_WINDOW', lambda: check(windows))

    windows.mainloop()
