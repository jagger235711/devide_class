"""
 -*- coding: utf-8 -*-

 @Time : 2021/10/27 15:21

 @Author : jagger

 @File : 阳光分班.py

 @Software: PyCharm 

 @contact: 252587809@qq.com

 -*- 功能说明 -*-

"""
import easygui as g

g.msgbox(msg='欢迎使用阳光分班程序！！！', title='你好!', ok_button='OK')
choice = g.buttonbox(msg='请选择你需要的服务内容: ', title='模式选择： ', choices=('生成人名单', '分班', '查看分班结果'), image=["28.png"])

if choice == '生成人名单':
    '''通过函数随机生成学生的信息，并保存到excel中生成一个information.xlsx文件，学生的每个信息包括
       学生的姓名，性别，年龄，语文成绩，数学成绩，英语成绩，总成绩，平均成绩
    '''
    import random
    import random as r
    import xlwt


    def GBK2312():  # 汉字国际扩展码
        head = r.randint(0xb0, 0xf7)
        body = r.randint(0xa1, 0xf9)  # 在head区号为55的那一块最后5个汉字是乱码,为了方便缩减下范围
        val = f'{head:x}{body:x}'
        st = bytes.fromhex(val).decode('gb2312')
        return st


    def first_name():  # 随机取姓氏字典
        first_name_list = [
            '赵', '钱', '孙', '李', '周', '吴', '郑', '王', '冯', '陈', '褚', '卫', '蒋', '沈', '韩', '杨', '朱', '秦', '尤', '许',
            '何', '吕', '施', '张', '孔', '曹', '严', '华', '金', '魏', '陶', '姜', '戚', '谢', '邹', '喻', '柏', '水', '窦', '章',
            '云', '苏', '潘', '葛', '奚', '范', '彭', '郎', '鲁', '韦', '昌', '马', '苗', '凤', '花', '方', '俞', '任', '袁', '柳',
            '酆', '鲍', '史', '唐', '费', '廉', '岑', '薛', '雷', '贺', '倪', '汤', '滕', '殷', '罗', '毕', '郝', '邬', '安', '常',
            '乐', '于', '时', '傅', '皮', '卞', '齐', '康', '伍', '余', '元', '卜', '顾', '孟', '平', '黄', '和', '穆', '萧', '尹',
            '姚', '邵', '堪', '汪', '祁', '毛', '禹', '狄', '米', '贝', '明', '臧', '计', '伏', '成', '戴', '谈', '宋', '茅', '庞',
            '熊', '纪', '舒', '屈', '项', '祝', '董', '梁']
        n = r.randint(0, len(first_name_list) - 1)
        f_name = first_name_list[n]
        return f_name


    def second_name():  # 随机获取名
        # 随机取数组中字符，取到空字符则没有second_name
        second_name_list = [GBK2312(), '']
        n = r.randint(0, 1)
        s_name = second_name_list[n]
        return s_name


    def last_name():  # 随机获取姓
        return GBK2312()


    def create_name():  # 生成随即名
        name = first_name() + second_name() + last_name()
        return name


    def getSex():  # 随机生成性别
        return random.choice(('男', '女'))


    def getAge():  # 随机生成年龄
        return str(r.randint(13, 15))


    def getChinese():  # 随机生成语文成绩
        return str(r.randint(0, 100))


    def getMath():  # 随机生成数学成绩
        return str(r.randint(0, 100))


    def getEnglish():  # 随机生成英语成绩
        return str(r.randint(0, 100))


    def main():  # 随机生成的学生信息保存到excel中
        work_book = xlwt.Workbook(encoding='utf-8')
        sheet = work_book.add_sheet('sheet1', cell_overwrite_ok=True)
        sheet.write(0, 0, '姓名')  # 写第零行的属性
        sheet.write(0, 1, '性别')
        sheet.write(0, 2, '年龄')
        sheet.write(0, 3, '语文成绩')
        sheet.write(0, 4, '数学成绩')
        sheet.write(0, 5, '英语成绩')
        sheet.write(0, 6, '总成绩')
        sheet.write(0, 7, '平均成绩')

        for i in range(1200):  # 将1200个数值插入excel中
            name = create_name()
            sex = getSex()
            age = getAge()
            chinese = getChinese()
            math = getMath()
            english = getEnglish()
            grade = str(int(chinese) + int(math) + int(english))
            avg = str(float(grade) // 3)
            sheet.write(i, 0, name)
            sheet.write(i, 1, sex)
            sheet.write(i, 2, age)
            sheet.write(i, 3, chinese)
            sheet.write(i, 4, math)
            sheet.write(i, 5, english)
            sheet.write(i, 6, grade)
            sheet.write(i, 7, avg)

        work_book.save('information.xls')


    if __name__ == '__main__':
        main()

elif choice == '分班':
    '''读取information.xlsx文件的数据，将学生的数据保存到一个字典列表中，对字典进行sort() 排序处理
    将对字典进行切片，插入到不同的班级表格中，从而达到分班的目的
    '''
    import pandas
    import pandas as pd

    dict = {}  # 定义一个字典
    data = pd.read_excel(r'D:\新建文件夹\PythonProject\小项目\大作业\information.xlsx')
    data.fillna("", inplace=True)
    list = []  # 定义一个数列

    for i in data.index.values:
        line = data.loc[i, ["姓名", "性别", "年龄", "语文成绩", "数学成绩", "英语成绩", "总成绩", "平均成绩"]].to_dict()
        list.append(line)
    dict = list  # 生成字典列表
    dict.sort(key=lambda s: int(s["平均成绩"]))

    writer = pd.ExcelWriter("分班结果.xlsx")

    for i in range(22):
        df = pandas.DataFrame(dict[i:1200:22])  # 对字典进行切片
        name = str(i + 1)  # 生成不同的表格
        df.to_excel(writer, sheet_name=name)  # 将数据插入表格中

    writer.save()
    writer.close()

elif choice == '查看分班结果':
    '''
    调用webbrowser 库 打开分班结果的xlsx文件
    '''
    import webbrowser

    webbrowser.open('分班结果.xlsx')
