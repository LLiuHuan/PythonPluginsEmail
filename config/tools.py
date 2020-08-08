# -*- coding: utf-8 -*-
import xlwt


def data_is_merge(xs,
                  lcol,
                  data,
                  start_row=2,
                  is_merge=False,
                  rcol=0,
                  is_lr=True,
                  is_bold=False):
    """
    按列插入数据
    :param xs: sheet对象
    :param col: 第n列
    :param data: 数据
    :param start_row: 起始列
    :param is_merge: 该列是否合并
    :return:
    """
    style = xlwt.XFStyle()  # Create Style
    if is_merge:
        new_col = list(set(data))
        new_col.sort(key=data.index)  # sort排序与原顺序一致

        alignment = xlwt.Alignment()  # 创建对其格式的对象 Create Alignment
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
        borders = xlwt.Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        font = xlwt.Font()
        font.bold = is_bold
        style.font = font
        style.borders = borders
        style.alignment = alignment

        if is_lr:
            for i, v in enumerate(new_col, start_row):
                count = data.count(v)
                end_row = start_row + count - 1
                xs.write_merge(start_row, end_row, lcol, rcol, v, style)
                start_row = end_row + 1
        else:
            for i, v in enumerate(new_col, start_row):
                xs.write(lcol, i, v, style)
    else:
        for i, v in enumerate(data, start_row):
            borders = xlwt.Borders()
            borders.left = 1
            borders.right = 1
            borders.top = 1
            borders.bottom = 1
            style.borders = borders
            xs.write(i, lcol, v, style)
    return xs


def change_data(data1, data2):
    """
    将列表数据转成xlwt需要的数据
    :param data1: 数据
    :param data2: 表头
    :return:
    """
    content = []
    for i, v in enumerate(data1):
        for ii, vv in enumerate(data2):
            if i == 0:
                content.append([v[vv]])
            else:
                content[ii].append(v[vv])
    return content