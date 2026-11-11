# -*- coding: utf-8 -*-
"""从报名/统计表里提炼"需跟进名单"。

很多真实表格最后一列是自由备注（"需重新提交""考核作业""报名表""待交"），
这一列才是班委真正要的行动项，但通用清洗引擎只会把它当备注保留、不会汇总。

用法：
  python followup.py <xlsx> --sheet Sheet1 --header-row 2 \
      --name-col 姓名 --id-col 学号 --remark-col 需重新提交 --out <目录>

输出：
  followup.md   一页"需跟进名单"，可直接发群
  followup.xlsx 可筛选的名单
"""
import argparse
import os
import sys
from collections import Counter
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

THIN = Side(style="thin", color="D5DBE2")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1B3A5C")
HEAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10)


def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--header-row", type=int, required=True)
    ap.add_argument("--name-col", default="姓名")
    ap.add_argument("--id-col", default="学号")
    ap.add_argument("--remark-col", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.input, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb[wb.sheetnames[0]]

    header = {}
    for c in range(1, ws.max_column + 1):
        h = cell_str(ws.cell(args.header_row, c).value)
        if h:
            header[h] = c
    if args.remark_col not in header:
        print("找不到备注列：%s ｜ 现有列：%s" % (args.remark_col, "、".join(header)))
        return 2

    c_name = header.get(args.name_col)
    c_id = header.get(args.id_col)
    c_rm = header[args.remark_col]

    todo, all_rows = [], 0
    for r in range(args.header_row + 1, ws.max_row + 1):
        name = cell_str(ws.cell(r, c_name).value) if c_name else ""
        sid = cell_str(ws.cell(r, c_id).value) if c_id else ""
        if not name and not sid:
            continue
        all_rows += 1
        remark = cell_str(ws.cell(r, c_rm).value)
        if remark:
            todo.append({"row": r, "name": name, "id": sid, "remark": remark})

    counts = Counter(t["remark"] for t in todo)
    os.makedirs(args.out, exist_ok=True)

    md = []
    md.append("# 需跟进名单：%s\n" % os.path.basename(args.input))
    md.append("> 生成时间：%s ｜ 来源行号为本表真实行号，可逐行回查。" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    md.append("> 判定口径：`%s` 列非空即为需跟进。\n" % args.remark_col)
    md.append("## 一、总览\n")
    md.append("| 指标 | 数值 |")
    md.append("| --- | ---: |")
    md.append("| 表中总人数 | %d 人 |" % all_rows)
    md.append("| **需跟进** | **%d 人（占 %.0f%%）** |" % (len(todo), len(todo) / max(1, all_rows) * 100))
    md.append("| 无需处理 | %d 人 |\n" % (all_rows - len(todo)))

    if counts:
        md.append("## 二、按原因分组\n")
        md.append("| 原因 | 人数 | 名单 |")
        md.append("| --- | ---: | --- |")
        for reason, n in counts.most_common():
            who = "、".join(t["name"] for t in todo if t["remark"] == reason)
            md.append("| %s | %d | %s |" % (reason, n, who))
        md.append("")

    md.append("## 三、逐人明细\n")
    md.append("| 原表行号 | 姓名 | 学号 | %s |" % args.remark_col)
    md.append("| ---: | --- | --- | --- |")
    for t in todo:
        md.append("| %d | %s | %s | %s |" % (t["row"], t["name"], t["id"], t["remark"]))
    md.append("")
    md.append("## 四、建议动作\n")
    md.append("1. 把上面名单直接发到群里，@ 到人，比逐个私聊快。")
    for i, (reason, n) in enumerate(counts.most_common(), start=2):
        md.append("%d. 「%s」%d 人，需要重新提交对应材料。" % (i, reason, n))
    md.append("")

    p = os.path.join(args.out, "followup.md")
    open(p, "w", encoding="utf-8").write("\n".join(md))

    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "需跟进"
    hdr = ["原表行号", "姓名", "学号", args.remark_col]
    ws2.append(hdr)
    for i in range(1, len(hdr) + 1):
        c = ws2.cell(1, i)
        c.font, c.fill, c.border = HEAD_FONT, HEAD_FILL, BORDER
        c.alignment = Alignment(horizontal="center")
    for t in todo:
        ws2.append([t["row"], t["name"], t["id"], t["remark"]])
        for i in range(1, len(hdr) + 1):
            ws2.cell(ws2.max_row, i).font = BODY_FONT
            ws2.cell(ws2.max_row, i).border = BORDER
    for i, w in enumerate([12, 14, 16, 20], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    wb2.save(os.path.join(args.out, "followup.xlsx"))

    print("总人数 %d ｜ 需跟进 %d 人 ｜ 分组：%s" % (
        all_rows, len(todo), "；".join("%s×%d" % (k, v) for k, v in counts.most_common())))
    print("产出：", os.path.join(args.out, "followup.md"), "、", os.path.join(args.out, "followup.xlsx"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
