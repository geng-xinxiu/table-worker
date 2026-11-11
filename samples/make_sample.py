#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_sample.py —— 生成"故意做脏"的班级报名样例表（用于验证表格工引擎）
=====================================================================

产物：samples/dirty_报名表.xlsx
要求：≥60 行数据、覆盖 10 类典型脏数据问题各 2–3 处、含标题行 + 合并单元格 + 两行表头。

脱敏声明：所有姓名均为化名，手机号为虚拟号（1380000xxxx），学号为虚构编号，
不含任何真实个人信息。

用法：
    python make_sample.py                 # 生成到 samples/dirty_报名表.xlsx
    python make_sample.py --out xxx.xlsx   # 指定输出文件
"""

import argparse
import random
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TITLE = "2025年秋季学期 班级报名信息汇总表（样例数据·化名与虚拟号）"
GROUP_ROW = ["学生基本信息", "学生基本信息", "学生基本信息",
             "组织信息", "组织信息",
             "成绩与提交", "成绩与提交", "成绩与提交", "成绩与提交"]
HEADER = ["名字", "学　号", "手机", "行政班", "院系", "所学专业", "平均绩点", "提交时间", "说明"]

CLASSES = ["2023级1班", "2023级2班", "2024级1班", "2024级2班", "2025级1班"]
COLLEGES = ["计算机学院", "机械工程学院", "经济与工商管理学院", "外国语学院", "建筑城规学院"]
MAJORS = ["计算机科学与技术", "软件工程", "机械设计制造及其自动化", "金融学", "英语",
          "建筑学", "数据科学与大数据技术", "工商管理"]
SURNAMES = ["李", "王", "张", "刘", "陈", "杨", "黄", "赵", "周", "吴",
            "徐", "孙", "马", "朱", "胡", "郭", "何", "林", "高", "罗"]
GIVENS = ["思远", "雨桐", "子航", "佳怡", "振宇", "梦琪", "浩然", "欣怡", "嘉豪", "雨欣",
          "文轩", "诗涵", "俊杰", "雅静", "一鸣", "书瑶", "博文", "晨曦", "锦程", "可欣",
          "明轩", "语桐", "天佑", "若曦", "泽宇"]

EXCEL_EPOCH = date(1899, 12, 30)


def excel_serial(day):
    return (day - EXCEL_EPOCH).days


def build_rows():
    """构造数据行，并按清单注入 10 类脏数据。返回 (rows, 注入说明)。"""
    rng = random.Random(20251109)
    rows = []
    used_names = set()

    def fake_name(seed):
        while True:
            name = SURNAMES[seed % len(SURNAMES)] + GIVENS[(seed * 7) % len(GIVENS)]
            if name not in used_names:
                used_names.add(name)
                return name
            seed += 1

    # ---- 50 行"基本干净"的数据（仍保留少量正常写法差异，属于真实世界的常态）
    for index in range(1, 51):
        rows.append([
            fake_name(index),
            "2023{:06d}".format(1000 + index),                  # 2023010010xx，11 位纯数字
            "1380000{:04d}".format(index),                      # 13800000001 … 虚拟号
            CLASSES[index % len(CLASSES)],
            COLLEGES[index % len(COLLEGES)],
            MAJORS[index % len(MAJORS)],
            round(rng.uniform(2.0, 4.5), 2),
            "2025-11-{:02d}".format(rng.randint(1, 9)),
            "",
        ])

    injections = []

    def add(note, row):
        rows.append(row)
        injections.append(note)

    # ---- R01 整行重复 3 处（完全照抄已有行）
    for source in (2, 15, 31):
        copied = list(rows[source - 1])
        add("R01 整行重复（照抄第 {} 行）".format(source), copied)

    # ---- R02 学号重复 3 处（不同姓名，同一学号 → 疑似重复提交）
    for offset, source in enumerate((4, 22, 44), start=1):
        row = list(rows[source - 1])
        row[0] = fake_name(900 + offset)
        add("R02 学号重复（{} 与第 {} 行重号）".format(row[1], source), row)

    # ---- R03 手机号重复 2 处（同名不同学号 → 同一人填了两次）
    for offset, source in enumerate((7, 38), start=1):
        row = list(rows[source - 1])
        row[1] = "2023{:06d}".format(2000 + offset)
        row[0] = fake_name(950 + offset)
        add("R03 手机号重复（{} 与第 {} 行同号）".format(row[2], source), row)

    # ---- R04 手机号格式异常 4 处
    bad_phone_rows = [
        ["1380000001", "少一位（10 位）"],
        ["138000000012", "多一位（12 位）"],
        ["138abcd0025", "混入字母"],
        ["138 0000 0033", "含空格（可自动修正）"],
    ]
    for offset, (phone, why) in enumerate(bad_phone_rows, start=1):
        add("R04 手机号异常（{}）".format(why), [
            fake_name(1000 + offset), "2023{:06d}".format(3000 + offset), phone,
            CLASSES[0], COLLEGES[0], MAJORS[0], 3.21, "2025/11/08", "手机号问题：{}".format(why)])
    # 全角数字手机号（可自动修正）
    add("R04 手机号异常（全角数字，可自动修正）", [
        fake_name(1010), "2023300105", "１３８０００００４２",
        CLASSES[1], COLLEGES[1], MAJORS[1], 3.35, "2025.11.8", "全角数字"])

    # ---- R05 学号格式异常 3 处
    for offset, (sid, why) in enumerate(
            [("2023A001", "字母在中间"), ("1234567", "位数过短（7 位）"), ("202330000010011", "位数过长（15 位）")],
            start=1):
        add("R05 学号异常（{}）".format(why), [
            fake_name(1100 + offset), sid, "1380000{:04d}".format(60 + offset),
            CLASSES[2], COLLEGES[2], MAJORS[2], 3.05, "11月7日", "学号问题：{}".format(why)])

    # ---- R06 必填项缺失 3 处（姓名空 / 手机号空 / 学号空）
    add("R06 必填缺失（姓名为空）", [
        None, "2023300201", "13800000071", CLASSES[3], COLLEGES[3], MAJORS[3], 2.88, "2025-11-06", "忘填姓名"])
    add("R06 必填缺失（手机号为空）", [
        fake_name(1201), "2023300202", None, CLASSES[4], COLLEGES[4], MAJORS[4], 3.66, "2025-11-06", "没留电话"])
    add("R06 必填缺失（学号为空）", [
        fake_name(1202), None, "13800000073", CLASSES[0], COLLEGES[0], MAJORS[5], 3.12, "2025-11-05", "学号漏填"])

    # ---- R07 姓名含空格/全角空格 5 处
    spaced = [("  {}  ".format(fake_name(1301)), "首尾半角空格"),
              ("{}　".format(fake_name(1302)), "尾部全角空格"),
              ("{}  {}".format(fake_name(1303)[0], fake_name(1303)[1:]), "姓名中间夹空格"),
              ("\u200b{}".format(fake_name(1304)), "前置零宽字符"),
              ("{} ".format(fake_name(1305)), "尾随空格")]
    for offset, (name, why) in enumerate(spaced, start=1):
        add("R07 姓名不规范（{}）".format(why), [
            name, "20233003{:02d}".format(offset), "1380000{:04d}".format(80 + offset),
            CLASSES[offset % len(CLASSES)], COLLEGES[offset % len(COLLEGES)], MAJORS[offset % len(MAJORS)],
            3.4, "2025-11-04", "姓名格式：{}".format(why)])

    # ---- R08 日期格式五花八门 6 处
    date_variants = [
        ("2025/11/9", "斜杠分隔"),
        ("11月9日", "中文月日（缺年份）"),
        ("2025.11.9", "点号分隔"),
        ("2025年11月9日", "中文年月日"),
        ("20251109", "紧凑 8 位"),
        (str(excel_serial(date(2025, 11, 9))), "Excel 日期序列号"),
    ]
    for offset, (value, why) in enumerate(date_variants, start=1):
        add("R08 日期格式（{}）".format(why), [
            fake_name(1401 + offset), "20233004{:02d}".format(offset), "1380000{:04d}".format(90 + offset),
            CLASSES[offset % len(CLASSES)], COLLEGES[offset % len(COLLEGES)], MAJORS[offset % len(MAJORS)],
            3.72, value, "日期写法：{}".format(why)])

    # ---- R10 数值列异常值 2 处
    add("R10 绩点异常（9.9 超上限）", [
        fake_name(1501), "2023300501", "13800000101", CLASSES[1], COLLEGES[1], MAJORS[1],
        9.9, "2025-11-03", "绩点填错"])
    add("R10 绩点异常（-0.5 低于下限）", [
        fake_name(1502), "2023300502", "13800000102", CLASSES[2], COLLEGES[2], MAJORS[2],
        -0.5, "2025-11-03", "绩点填错"])

    # ---- R11 白名单不匹配 3 处
    add("R11 班级不在白名单（2025级3班）", [
        fake_name(1601), "2023300601", "13800000111", "2025级3班", COLLEGES[0], MAJORS[0],
        3.58, "2025-11-02", "班级名写错"])
    add("R11 学院名称不规范（计算机科学与技术学院）", [
        fake_name(1602), "2023300602", "13800000112", CLASSES[1], "计算机科学与技术学院", MAJORS[1],
        3.44, "2025-11-02", "学院名写全称"])
    add("R11 学院名称不规范（计院）", [
        fake_name(1603), "2023300603", "13800000113", CLASSES[2], "计院", MAJORS[2],
        3.31, "2025-11-02", "学院名写简称"])

    # ---- 穿插一个整行空白行（应被跳过，不计入数据行）
    rows.insert(25, [None] * len(HEADER))
    return rows, injections


def build_workbook(output_path):
    rows, injections = build_rows()
    wb = Workbook()
    ws = wb.active
    ws.title = "报名表"

    width_map = {"A": 14, "B": 16, "C": 17, "D": 13, "E": 22, "F": 26, "G": 10, "H": 16, "I": 26}
    for letter, width in width_map.items():
        ws.column_dimensions[letter].width = width

    # 第 1 行：标题（跨全部列合并）
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADER))
    title_cell = ws.cell(row=1, column=1, value=TITLE)
    title_cell.font = Font(bold=True, size=13)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 第 2 行：分组表头（合并单元格）—— 模拟"列名被合并单元格搞乱"
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)
    ws.merge_cells(start_row=2, start_column=4, end_row=2, end_column=5)
    ws.merge_cells(start_row=2, start_column=6, end_row=2, end_column=9)
    for column in (1, 4, 6):
        cell = ws.cell(row=2, column=column, value=GROUP_ROW[column - 1])
        cell.font = Font(bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="D9E1F2")

    # 第 3 行：真实表头（用同义词，测试列名规范化）
    header_fill = PatternFill("solid", fgColor="1F3864")
    for position, header in enumerate(HEADER, start=1):
        cell = ws.cell(row=3, column=position, value=header)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 数据行
    for row_offset, row in enumerate(rows, start=4):
        for column_offset, value in enumerate(row, start=1):
            ws.cell(row=row_offset, column=column_offset, value=value)

    # 尾部再留 2 行空白（测试容错）
    last_row = 4 + len(rows)
    ws.cell(row=last_row + 1, column=1, value=None)
    ws.cell(row=last_row + 2, column=1, value=None)
    ws.freeze_panes = "A4"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
    return rows, injections


def main(argv=None):
    parser = argparse.ArgumentParser(description="生成故意做脏的班级报名样例表")
    parser.add_argument("--out", default=str(Path(__file__).parent / "samples" / "dirty_报名表.xlsx"),
                        help="输出 xlsx 路径")
    args = parser.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    output = Path(args.out)
    rows, injections = build_workbook(output)
    print("已生成样例表：{}".format(output.resolve()))
    print("工作表行结构：第 1 行标题（合并 9 列）+ 第 2 行分组表头（3 处合并）+ 第 3 行真实表头 + 数据行")
    print("数据行数：{} 行（含 1 行整行空白）".format(len(rows)))
    print("注入的脏数据：{} 处".format(len(injections)))
    for index, note in enumerate(injections, 1):
        print("  {:>2}. {}".format(index, note))
    print("\n脱敏声明：姓名为化名，手机号为虚拟号 1380000xxxx，学号为虚构编号。")


if __name__ == "__main__":
    main()
