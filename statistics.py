#! /usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import division

import os
import sys
from collections import OrderedDict
import itertools
import pandas as pd


OUTPUT_FILE_NAME = 'output.xlsx'

help_text = """Please read this help:
第一个参数为Excel所在的文件夹
第二个参数为 part1 / part2 / part3 / part4 中的一个, 如果不传则统计所有部分
"""


def init_an_order_dict(a_list, init_value=0):
    return OrderedDict().fromkeys(a_list, value=init_value)


def get_file():
    """
    获取所有要操作的Excel文件
    """
    if len(sys.argv) < 2:
        print(help_text)
        sys.exit(1)
    excel_dir = sys.argv[1]
    files_in_dir = os.listdir(excel_dir)
    return [
        os.path.join(excel_dir, f) for f in files_in_dir if '.xls' in f and OUTPUT_FILE_NAME not in f
    ]  # .xls / .xlsx


def part_1_1():
    part1_1 = init_an_order_dict(year_list_pured)
    for year in year_list:
        part1_1[year] += 1
    return part1_1


def part_1_2():
    """
    每年被引次数
    """
    part1_2 = init_an_order_dict(year_list_pured)
    for i in xrange(0, len(year_list)):
        part1_2[year_list[i]] += int(col_bei_yin_zheng_ci_shu[i])
    return part1_2


def part_1_3():
    part1_3 = init_an_order_dict(year_list_pured)
    for i in xrange(0, len(year_list)):
        patent_str = str(col_bei_yin_zheng_zhuan_li[i]).strip()
        if patent_str != "nan":
            patent_list = patent_str.split('; ')
            for p in patent_list:
                if p in set(col_gong_kai_hao):
                    part1_3[year_list[i]] += 1
        else:
            pass
    return part1_3


def part_1_4():
    part1_4 = init_an_order_dict(year_list_pured, init_value="")
    for i in xrange(0, len(year_list)):
        if ';' in col_shen_qing_ren[i]:
            part1_4[year_list[i]] += "%s " % (col_xu_hao[i])
    return part1_4


def part_1_5():
    part1_5 = init_an_order_dict(year_list_pured, init_value="")
    for i in xrange(0, len(year_list)):
        if ';' in col_shen_qing_ren[i] and u'大学' in col_shen_qing_ren[i]:
            part1_5[year_list[i]] += "%s " % (col_xu_hao[i])
    return part1_5


def part1():
    part1_1 = part_1_1()
    part1_2 = part_1_2()
    part1_3 = part_1_3()
    part1_4 = part_1_4()
    part1_5 = part_1_5()
    keys = part1_1.keys()
    output = []
    for k in keys:
        line = [
            company_code, k, part1_1[k], part1_2[k], part1_3[k],
            u"%s(%s个)" % (part1_4[k].strip(), 0 if part1_4[k].strip() == u'' else len(part1_4[k].strip().split(' '))),
            u"%s(%s个)" % (part1_5[k].strip(), 0 if part1_5[k].strip() == u'' else len(part1_5[k].strip().split(' ')))
        ]
        output.append(line)
    return output


def part2():
    patent_ipc_combinations = init_an_order_dict(year_list_pured, init_value=None)
    part2_creation = init_an_order_dict(year_list_pured)
    part2_reuse = init_an_order_dict(year_list_pured)
    ipc_combination_list = []
    for line_no in xrange(0, len(year_list)):  # 遍历Excel表格，统计每行的IPC组合并按年累加
        ipc_str = str(col_ipc[line_no]) if col_ipc[line_no] != 'nan' else ''
        ipc_list = [ipc[0:4] for ipc in ipc_str.split('; ')]  # 取前4个字符
        ipc_set = set(ipc_list)  # 每个专利去重后的IPC
        combinations = [i for i in itertools.combinations(ipc_set, 2)]  # Cn2  每个专利的IPC组合
        ipc_combination_list.append(combinations)
        year = year_list[line_no]
        if patent_ipc_combinations[year] is None:
            patent_ipc_combinations[year] = []
        else:
            patent_ipc_combinations[year] += combinations

    for y in year_list_pured:
        n_years_before_data = []  # 5年窗口期所有组合
        if year_list_pured.index(y) == 0:
            pass  # 第一年不统计
        elif int(year_list_pured[0]) + 5 <= int(y):  # 前面有多于5年只算前5年, 5年是看数不是看个数
            for i in range(1, 6):
                if y-i in year_list_pured:
                    n_years_before_data += patent_ipc_combinations[y - i]
        else:  # 前面没有5年的有几年算几年
            for j in range(0, year_list_pured.index(y) + 1):
                if y-j in year_list_pured:
                    n_years_before_data += patent_ipc_combinations[y - j]

        new_combination = [c1 for c1 in patent_ipc_combinations[y] if c1 not in n_years_before_data]  # 当年新增的组合
        repeated_combination = [c2 for c2 in patent_ipc_combinations[y] if c2 in n_years_before_data]  # 旧组合

        creation_top_value = len(new_combination)  # 分子: 新增年份的新组合数
        creation_bottom_value = len(patent_ipc_combinations[y])  # 分母：新增年份的组合数
        # 如果分子或分母有一个是0则值为0
        part2_creation[y] = '0' if creation_bottom_value == 0 or creation_bottom_value == 0 else "=%s/%s" % (
            creation_top_value, creation_bottom_value)

        reuse_top_value = 0  # 分子：旧的组合对应的专利数
        for patent in ipc_combination_list:
            if set(patent) & set(repeated_combination):
                reuse_top_value += 1

        reuse_bottom_value = creation_bottom_value  # 分母：新增年份的组合数, 和 creation 一样

        part2_reuse[y] = '0' if reuse_bottom_value == 0 or reuse_top_value == 0 else "=%s/%s" % (
            reuse_top_value, reuse_bottom_value)

    return [
        [company_code, str(nian), part2_creation[nian], part2_reuse[nian]] for nian in year_list_pured
    ]


def part3():
    output_list = list()
    for i in xrange(0, col_gong_kai_hao.count()):
        col_d_value = col_gong_kai_hao[i]
        if unicode(col_yin_zheng_zhuan_li[i]) != u'nan':
            for r in col_yin_zheng_zhuan_li[i].split('; '):
                output_list.append(
                    (company_code, i+1, r, col_d_value)
                )
        if unicode(col_bei_yin_zheng_zhuan_li[i]) != u'nan':
            for s in col_bei_yin_zheng_zhuan_li[i].split('; '):
                output_list.append(
                    (company_code, i+1, col_d_value, s)
                )
    return output_list


def part4():
    output_list = list()
    for i in xrange(0, col_shen_qing_ren.count()):
        col_i_value = col_shen_qing_ren[i]  # 申请人还有多个人的情况
        applyer_list = col_i_value.split('; ')
        if unicode(col_yin_zheng_shen_qing_ren[i]) != 'nan':
            for t in col_yin_zheng_shen_qing_ren[i].split('; '):
                for applyer1 in applyer_list:
                    output_list.append(
                        (company_code, i+1, t, applyer1)
                    )
        if unicode(col_bei_yin_zheng_shen_qing_ren[i]) != 'nan':
            for u in col_bei_yin_zheng_shen_qing_ren[i].split('; '):
                for applyer2 in applyer_list:
                    output_list.append(
                        (company_code, i+1, applyer2, u)
                    )
    return output_list


def write_part1():
    try:
        result1 = pd.concat(part1_to_write)
        result1.to_excel(
            writer, sheet_name='part1', index=False,
            header=[u'公司', u'年份', u'专利总数', u'被引次数', u'被引且出现在D列', u'申请人数大于2', u'申请人数大于2且含大学']
        )
    except:
        print "part1 error, pass..."


def write_part2():
    try:
        result2 = pd.concat(part2_to_write)
        result2.to_excel(writer, sheet_name='part2', header=[u'公司', u'年份', u'creation', u'reuse'], index=False)
    except:
        print "part2 error, pass..."


def write_part3():
    try:
        result3 = pd.concat(part3_to_write)
        result3.to_excel(writer, sheet_name='part3', header=[u'公司', u'序号', u'A', u'B'], index=False)
    except:
        print "part3 error, pass..."


def write_part4():
    try:
        result4 = pd.concat(part4_to_write)
        result4.to_excel(writer, sheet_name='part4', header=[u'公司', u'序号', u'A', u'B'], index=False)
    except:
        print "part4 error, pass..."


if __name__ == '__main__':
    if len(sys.argv) == 2:
        part_arg = 'all'  # 不传part参数则统计所有部分
    elif len(sys.argv) == 3:
        part_arg = sys.argv[2]
    else:
        print(help_text)
        sys.exit(1)

    files = get_file()
    writer = pd.ExcelWriter(OUTPUT_FILE_NAME, engine='xlsxwriter')
    part1_to_write = []
    part2_to_write = []
    part3_to_write = []
    part4_to_write = []

    for xls in files:
        base_name = os.path.basename(xls)
        company_code = str(base_name.split('.')[0])
        print '\n', company_code

        df = pd.read_excel(xls)

        col_shen_qing_ri = df[u'申请日']
        col_shen_qing_ren = df[u'申请人']  # I列
        col_xu_hao = df[u'序号']
        col_gong_kai_hao = df[u'公开（公告）号']  # D列
        col_yin_zheng_zhuan_li = df[u'引证专利']  # R列
        col_ipc = df['IPC']  # P列
        col_bei_yin_zheng_zhuan_li = df[u'被引证专利']  # S列
        col_bei_yin_zheng_ci_shu = df[u'被引证次数']
        col_yin_zheng_shen_qing_ren = df[u'引证申请人']  # T
        col_bei_yin_zheng_shen_qing_ren = df[u'被引证申请人']  # U

        year_list = [y.year for y in col_shen_qing_ri]
        year_list_pured = sorted(set(year_list))

        if part_arg == 'part1':
            part1_to_write.append(pd.DataFrame(part1()))
        elif part_arg == 'part2':
            part2_to_write.append(pd.DataFrame(part2()))
        elif part_arg == 'part3':
            part3_to_write.append(pd.DataFrame(part3()))
        elif part_arg == 'part4':
            part4_to_write.append(pd.DataFrame(part4()))
        elif part_arg == 'all':
            part1_to_write.append(pd.DataFrame(part1()))
            part2_to_write.append(pd.DataFrame(part2()))
            part3_to_write.append(pd.DataFrame(part3()))
            part4_to_write.append(pd.DataFrame(part4()))
        else:
            print(help_text)
            sys.exit(1)

    if part_arg == 'part1':
        write_part1()
    elif part_arg == 'part2':
        write_part2()
    elif part_arg == 'part3':
        write_part3()
    elif part_arg == 'part4':
        write_part4()
    else:
        write_part1()
        write_part2()
        write_part3()
        write_part4()

    writer.save()
    print("\nDone! 结果已输出到 %s 中\n" % OUTPUT_FILE_NAME)
