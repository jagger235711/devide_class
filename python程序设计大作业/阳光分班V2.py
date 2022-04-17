"""
 -*- coding: utf-8 -*-

 @Time : 2021/10/27 15:29

 @Author : jagger

 @File : 阳光分班V2.py

 @Software: PyCharm 

 @contact: 252587809@qq.com

 -*- 功能说明 -*-
配置以文件的形式保存
窗体使用tk
名字性别随机产生
1,'生'成人员名单
2,'图'形化界面

"""
import random
import xlwt
import easygui


class divide_classes:
    def __init__(self):
        self.excelName = '分班结果.xls'
        # self.stuList = [['姓名'], ['性别'], ['是否完成分班']]
        self.stuList = []
        self.classList = []
        self.firstNameList = [
            '赵', '钱', '孙', '李', '周', '吴', '郑', '王', '冯', '陈', '褚', '卫', '蒋', '沈', '韩', '杨', '朱', '秦', '尤', '许',
            '何', '吕', '施', '张', '孔', '曹', '严', '华', '金', '魏', '陶', '姜', '戚', '谢', '邹', '喻', '柏', '水', '窦', '章',
            '云', '苏', '潘', '葛', '奚', '范', '彭', '郎', '鲁', '韦', '昌', '马', '苗', '凤', '花', '方', '俞', '任', '袁', '柳',
            '酆', '鲍', '史', '唐', '费', '廉', '岑', '薛', '雷', '贺', '倪', '汤', '滕', '殷', '罗', '毕', '郝', '邬', '安', '常',
            '乐', '于', '时', '傅', '皮', '卞', '齐', '康', '伍', '余', '元', '卜', '顾', '孟', '平', '黄', '和', '穆', '萧', '尹',
            '姚', '邵', '堪', '汪', '祁', '毛', '禹', '狄', '米', '贝', '明', '臧', '计', '伏', '成', '戴', '谈', '宋', '茅', '庞',
            '熊', '纪', '舒', '屈', '项', '祝', '董', '梁']
        self.secondNameList = [
            '俊', '威', '英', '健', '壮', '焕', '挺', '帅', '秀', '伟', '武', '雄', '巍', '松', '柏', '山', '石', '婢', '娟', '姣', '妯',
            '姿', '媚', '婉', '丽', '妖', '美', '倩', '兰', '颖', '灵', '睿', '锐', '哲', '慧', '敦', '迪', '明', '晓', '显', '悉', '晰',
            '维', '学', '思', '悟', ' 析', '文', '书', '勤', '蔼', '仁', '容', '德', '轩', '贤', '良', '伦', '正', '清', '义', '诚', '直',
            '道', '达', '耀', '兴', '荣', '华', '旺', '盈', '丰', '余', '昌', '盛', '安', '静', '顺', '通', '坦', '泰', '然', '宁', '定',
            '和', '康', '毅', '独', '刚', '强', '衡', '韧', '坚', '力', '决', '定', '立', '主', '志', '意', '自'
        ]

    def produce_stuList(self):
        '''
        生成学生信息
        数据结构为【{}】列表套字典 即用列表储存学生 字典中存放的是学生的信息：名字、性别、是否分了班
        Returns:

        '''
        for i in range(1200):
            _dict = {}  # 列表套字典
            _dict['name'] = random.choice(self.firstNameList) + random.choice(self.secondNameList)
            _dict['sex'] = random.choice(['男', '女'])
            # _dict['flag'] = False  # 是否完成分班
            self.stuList.append(_dict)

    def allocation(self):
        '''
        将学生分配到班级中去
        数据结构为【【{}】】 列表套列表再套字典 最外层列表放的是班级共二十二个。第二层列表放的是学生，每班大约1200/22人。字典存放的是学生信息
        classList[_class[stu{}]]
        Returns:

        '''
        for i in range(22):
            _class = []
            stuList = self.stuList
            for j in range(1200 // 22):
                stu = random.choice(stuList)
                _class.append(stu)
                stuList.remove(stu)
            self.classList.append(_class)
        for i in stuList:
            __class = random.choice(self.classList)
            __class.append(i)

    def writer(self):
        '''
        将分班结果写入文件
        Returns:

        '''
        work_book = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        j = 1
        for i in self.classList:

            table = work_book.add_sheet('%s 班' % j)  # 添加工作表
            j += 1
            # 向工作表写入数据
            table.write(0, 0, '姓名:')  # 写第零行的属性
            table.write(0, 1, '性别:')
            k = 1
            for stu in i:
                table.write(k, 0, stu['name'])
                table.write(k, 1, stu['sex'])
                # table.write(k, 2, stu['flag'])
                k += 1

        work_book.save(self.excelName)  # 保存工作簿

    # def get_excelName(self):
    #     return self.excelName
    #
    # def set_excelName(self, name):
    #     self.excelName = name


def main():
    if easygui.ccbox(msg='欢迎来到分班小程序！', choices=['开始分班', '退出分班']):
        try:
            li = divide_classes()
            li.produce_stuList()
            li.allocation()
            li.writer()
        except Exception as e:
            easygui.msgbox('程序出现错误！' + str(e))
        else:
            easygui.msgbox('分班成功！')
    else:
        easygui.msgbox('拜拜！欢迎下次再来')


if __name__ == '__main__':
    main()
