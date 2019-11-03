import pandas as pd


def get_dict_class_level(path='lib/上课信息表.xlsx'):
    '''
    获取优先级字典
    :param path:
    :return:
    '''
    df = pd.read_excel(io=path, sheet_name='优先级')
    ls_class_name = df['课程名称'].tolist()
    ls_class_level = df['优先级'].tolist()
    dict_class_level = dict(zip(ls_class_name, ls_class_level))
    return dict_class_level


def get_dict_classroom(path='lib/上课信息表.xlsx', name='开放教室'):
    '''
    获取教室字典
    :param path:
    :param name:
    :return:
    '''
    df = pd.read_excel(io=path, sheet_name=name)
    ls_class_num = df['教室'].tolist()
    ls_class_vol = df['容量'].tolist()
    dict_class_room = dict(zip(ls_class_num, ls_class_vol))
    return dict_class_room


def get_dict_date(path='lib/上课信息表.xlsx', name='开放上课时间'):
    '''
    获取上课日期
    :param path:
    :param name:
    :return:
    '''
    df = pd.read_excel(io=path, sheet_name=name)
    ls_time_part = df['时间段'].tolist()
    ls_date = df['日期'].tolist()
    dict_date = dict(zip(ls_time_part, ls_date))
    return dict_date


def get_dict_time(path='lib/上课信息表.xlsx', name='开放上课时间'):
    '''
    获取上课日期
    :param path:
    :param name:
    :return:
    '''
    df = pd.read_excel(io=path, sheet_name=name)
    ls_time_part = df['时间段'].tolist()
    ls_time = df['时间'].tolist()
    dict_time = dict(zip(ls_time_part, ls_time))
    return dict_time
# dict_class_level=get_dict_class_level()
# dict_classroom = get_dict_classroom()
# dict_date=get_dict_date()
# dict_time=get_dict_time()
