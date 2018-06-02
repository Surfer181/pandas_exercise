#! /usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals, division, print_function

import os
import sys
from pprint import pprint
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
    return [f for f in files_in_dir if '.xls' in f and OUTPUT_FILE_NAME not in f]  # .xls / .xlsx


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
        if ';' in col_shen_qing_ren[i] and '大学' in col_shen_qing_ren[i]:
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
            "%s(%s个)" % (part1_4[k].strip(), 0 if part1_4[k].strip() == u'' else len(part1_4[k].strip().split(' '))),
            "%s(%s个)" % (part1_5[k].strip(), 0 if part1_5[k].strip() == u'' else len(part1_5[k].strip().split(' ')))
        ]
        output.append(line)
    return output


def part3():
    output_list = list()
    for i in xrange(0, col_gong_kai_hao.count()):
        col_d_value = col_gong_kai_hao[i]
        if unicode(col_yin_zheng_zhuan_li[i]) != 'nan':
            for r in col_yin_zheng_zhuan_li[i].split('; '):
                output_list.append(
                    (company_code, i+1, r, col_d_value)
                )
        if unicode(col_bei_yin_zheng_zhuan_li[i]) != 'nan':
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
    data_frame_output1 = pd.DataFrame(part1())
    data_frame_output1.to_excel(
        writer, sheet_name='part1', index=False,
        header=['公司', '年份', '专利总数', '被引次数', '被引且出现在D列', '申请人数大于2', '申请人数大于2且含大学']
    )


def write_part2():
    pass


def write_part3():
    data_frame_output3 = pd.DataFrame(part3())
    data_frame_output3.to_excel(writer, sheet_name='part3', header=['公司', '序号', 'A', 'B'], index=False)


def write_part4():
    data_frame_output4 = pd.DataFrame(part4())
    data_frame_output4.to_excel(writer, sheet_name='part4', header=['公司', '序号', 'A', 'B'], index=False)


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

    for xls in files:
        base_name = os.path.basename(xls)
        company_code = str(base_name.split('.')[0])
        print(company_code)

        df = pd.read_excel(xls)

        col_shen_qing_ri = df['申请日']
        col_shen_qing_ren = df['申请人']  # I列
        col_xu_hao = df['序号']
        col_gong_kai_hao = df['公开（公告）号']  # D列
        col_yin_zheng_zhuan_li = df['引证专利']  # R列
        col_bei_yin_zheng_zhuan_li = df['被引证专利']  # S列
        col_bei_yin_zheng_ci_shu = df['被引证次数']
        col_yin_zheng_shen_qing_ren = df['引证申请人']  # T
        col_bei_yin_zheng_shen_qing_ren = df['被引证申请人']  # U

        year_list = [y.year for y in col_shen_qing_ri]
        year_list_pured = sorted(set(year_list))

        if part_arg == 'part1':
            write_part1()
        elif part_arg == 'part2':
            write_part2()
        elif part_arg == 'part3':
            write_part3()
        elif part_arg == 'part4':
            write_part4()
        elif part_arg == 'all':
            write_part1()
            write_part2()
            write_part3()
            write_part4()
        else:
            print(help_text)
            sys.exit(1)

    writer.save()
    print("\nDone! 结果已输出到 %s 中\n" % OUTPUT_FILE_NAME)