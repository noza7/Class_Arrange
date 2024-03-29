# import pandas as pd
from lib.func import *

# 打开表格
path = '2019-2020第一学期开放教育开课一览表.xlsx'
# df = pd.read_excel(io=path, sheet_name='自开开课一览表')
df = pd.read_excel(io=path, sheet_name='网授开课一览表')
col_names_ls = df.columns.tolist()  # 获取标题行列表
col_names_ls.insert(0, '时间段')
# col_names_ls.insert(1, '教室')

# todo 提取数据结构

# todo 唯一值列表
ls_teachers = df['任课教师'].unique().tolist()
ls_courses = df['课程名称'].unique().tolist()
ls_classes = df['专业班级'].unique().tolist()


# todo 建立字典


def get_courses_dict():
    '''
    以课程为键，建立"课程:班级"字典
    :return: 返回"课程:班级"字典
    '''
    courses_dict = dict()  # 创建课程空字典
    for course in ls_courses:  # 从课程列表中取课程
        ls = []  # 创建一个空列表，用来存放班级
        for i in range(df.shape[0]):  # 从表中提取数据
            if df.iloc[i, 0] == course:  # 如果找到相同的课程名称
                d = dict()  # 创建一个空字典用来存放"班级:其它"信息字典

                # 把其它信息作为一个列表
                data_ls = []
                for j in range(2, df.shape[1]):
                    data_ls.append(df.iloc[i, j])

                d.update({df.iloc[i, 1]: data_ls})  # 加入字典
                # 把班级----其他信息字典加作为元素入列表
                ls.append(d)  # 把对应的专业班级添加进列表

        courses_dict.update({course: ls})  # 得到"课程:班级"字典
    return courses_dict


# todo 思路：通过课程，找到管理班（包含任课教师等所有信息）;
# todo 从课程列表取出一门课;
# todo 该课程对应的值也就是班级列表中的每一个班在班级列表（总表）中进行查询;
# todo 如果每个值都找到了，把数据以列表的形式写入结果列表;
# todo 同时从课程总列表中弹出这个课程，从班级列表（总表）中弹出找到的班级;


def get_1_time(courses_dict, set_classes, ls_courses, set_teachers, CLASSROOM_NUM, dict_class_level):
    '''
    # 获取一个时间段数据
    :param courses_dict: 课程所对应的班级字典列表
    :param set_classes: 班级集合（总表）
    :param ls_courses: 课程列表（每一次排课需要传入刨除已成功的课程）
    :return:data_ls(一次排课数据), have_courses_set(已被安排的课程)
    '''
    # 结果数据列表
    data_ls = []
    # 已选课程集合
    have_courses_set = set()
    # 用来记录课程名称
    classes_ls_ = []
    # 初始化计数器
    COUNTER = 1
    # 从优先级课程字典取出按顺序取出课程
    for course, _ in dict_class_level.items():
        # print(course)
        if not COUNTER == 0:
            classes_ls = courses_dict[course]  # 课程所对应的班级字典列表
            # 获取班级名称集合
            classname_set = set()
            # 获取任课教师集合
            tescher_set = set()
            for class_dict in classes_ls:
                classname = list(class_dict.keys())[0]
                classname_set.add(classname)
                for k, v in class_dict.items():
                    tescher_set.add(v[0])

            # 用集合判断班级名称是否在班级列表（总表）中
            if classname_set < set_classes and tescher_set < set_teachers:
                have_courses_set.add(course)
                for class_dict in classes_ls:
                    for k, v in class_dict.items():
                        ls = []
                        ls.append(course)  # 添加课程名称
                        ls.append(k)  # 添加班级名称
                        for i in v:
                            ls.append(i)  # 添加其他信息
                        data_ls.append(ls)
                        # 记录已编排课程名称
                        classes_ls_.append(course)
                        # 列表去重
                        classes_ls_ = list(set(classes_ls_))
                        # print(classes_ls_)
                        # print(len(data_all_ls))
                set_classes = set_classes - classname_set
                set_teachers = set_teachers - tescher_set
                # dict_class_level.pop(course)  # 从优先级课程列表弹出课程
        else:
            print('教室满了')
            break
        COUNTER = len(classes_ls_) % CLASSROOM_NUM
        # print(classes_ls_)
        # print(COUNTER)
    return data_ls, have_courses_set


# 获取班级、课程名称、任课教师集合
set_classes = set(ls_classes)
set_courses = set(ls_courses)
set_teachers = set(ls_teachers)
ls_classes_ = ls_classes.copy()  # 用深度拷贝，是复制不是指针
courses_dict = get_courses_dict()  # 调用函数创建"课程:班级"字典

# todo 程序执行部分
# 数据初始化
i = 1
data_all_ls = []
dict_class_level = get_dict_class_level()  # 获取优先级字典
dict_classroom = list(get_dict_classroom().keys())  # 获取教室字典
dict_date = get_dict_date()  # 获取上课日期
dict_time = get_dict_time()  # 获取上课时间
while len(set_courses) > 0:
    # print(len(set_courses))  # 如果课程长度大于零，说明课程没有被排完
    # todo 设置教室数量
    CLASSROOM_NUM = 14

    ls_courses = list(set_courses)
    # 获取排课列表
    data_ls, have_courses_set = get_1_time(courses_dict, set_classes, ls_courses, set_teachers, CLASSROOM_NUM,
                                           dict_class_level)
    # 添加时间段序号,及教室号
    class_room_num = 0
    for data in data_ls:
        data.insert(0, str(i))
        # try:
        #     data.insert(1, list(dict_kf_classroom.keys())[class_room_num])
        #     # print(data)
        # except:
        #     break
        data_all_ls.append(data)
        class_room_num += 1
    # 从课程中减去已排课程
    set_courses = set_courses - have_courses_set
    # 从优先级课程字典中弹出已排课程
    for have_course in list(have_courses_set):
        dict_class_level.pop(have_course)
    i += 1

# 插入教室
j = 0
data_all_ls[0].insert(1, dict_classroom[0])
for i in range(1, len(data_all_ls)):
    if data_all_ls[i][1] == data_all_ls[i - 1][2]:
        data_all_ls[i].insert(1, dict_classroom[j % len(dict_classroom)])
    else:
        j += 1
        data_all_ls[i].insert(1, dict_classroom[j % len(dict_classroom)])
col_names_ls.insert(1, '教室')

# 插入日期、时间
for i in range(len(data_all_ls)):
    t_key = data_all_ls[i][0]  # 总表时间段作为键值
    data_all_ls[i].insert(2, dict_date[int(t_key)])
    data_all_ls[i].insert(3, dict_time[int(t_key)])
col_names_ls.insert(2, '日期')
col_names_ls.insert(3, '时间')
# todo 写入Excel
data_all = pd.DataFrame(data_all_ls, columns=col_names_ls)
writer = pd.ExcelWriter('排课结果.xlsx')
data_all.to_excel(writer, '排课结果', index=False)
writer.save()
