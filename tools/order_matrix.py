# -*- coding: utf-8 -*-
"""团购订单矩阵 → 对账单

输入是班委常见的"转置表"：A 列是商品/属性名，从第 2 列起每一列是一个人的订单。
输出：
  1) orders_clean.xlsx  每人一行（姓名/订单类型/应付金额/支付状态/下单时间/件数/商品清单）
  2) items_matrix.xlsx  商品 × 人 的原始购买矩阵（转置后可读版）
  3) reconcile.md       一页对账单：谁买了什么、谁还没付、金额合计、数据异常
  4) run.log            机器日志（毫秒时间戳）

只做本地计算，不联网、不上传。所有数字来自本文件真实计算。
"""
import argparse
import json
import os
import re
import sys
import time
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

META_KEYS = ("商品", "订单类型", "应付金额", "支付状态", "下单时间")

THIN = Side(style="thin", color="D5DBE2")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1B3A5C")
WARN_FILL = PatternFill("solid", fgColor="FBE9E7")
HEAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10)


def parse_amount(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("¥", "").replace("￥", "").replace(",", "").replace("元", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_time(v):
    if v is None:
        return None, "缺失"
    if isinstance(v, datetime):
        return v, ""
    s = str(v).strip()
    patterns = [
        (r"^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$", "YMD"),
        (r"^(\d{1,2})\.(\d{1,2})\s+(\d{1,2}):(\d{2})$", "MD"),
    ]
    for pat, kind in patterns:
        m = re.match(pat, s)
        if m:
            g = m.groups()
            if kind == "YMD":
                return datetime(int(g[1 - 1]), int(g[1]), int(g[2]), int(g[3]), int(g[4])), ""
            return datetime(2026, int(g[0]), int(g[1]), int(g[2]), int(g[3])), "缺年份，按 2026 年补全"
    return None, "无法解析：%s" % s


def main():
    ap = argparse.ArgumentParser(description="团购订单矩阵 → 对账单")
    ap.add_argument("input")
    ap.add_argument("--out", required=True)
    ap.add_argument("--sheet", default=None)
    args = ap.parse_args()

    t0 = time.perf_counter()
    out = args.out
    os.makedirs(out, exist_ok=True)
    log_lines = []

    def log(msg):
        line = "[%s] %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3], msg)
        log_lines.append(line)
        print(line)

    wb = openpyxl.load_workbook(args.input, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb[wb.sheetnames[0]]
    log("读取 %s ｜ 工作表=%s ｜ %d 行 x %d 列" % (os.path.basename(args.input), ws.title, ws.max_row, ws.max_column))

    # ---------- 1. 解析矩阵 ----------
    row_labels = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
    people = []
    for c in range(2, ws.max_column + 1):
        name = ws.cell(1, c).value
        if name and str(name).strip():
            people.append({"col": c, "name": str(name).strip()})
    log("识别到 %d 位下单人：%s" % (len(people), "、".join(p["name"] for p in people)))

    meta_rows = {}
    item_rows = []
    for idx, label in enumerate(row_labels, start=1):
        if label is None:
            continue
        label = str(label).strip()
        if idx == 1:
            continue
        if label in META_KEYS:
            meta_rows[label] = idx
        else:
            item_rows.append((idx, label))
    log("属性行：%s ｜ 商品行：%d 条" % ("、".join(meta_rows) or "无", len(item_rows)))

    # 重复商品行检测
    seen = {}
    dup_items = []
    for r, label in item_rows:
        if label in seen:
            dup_items.append((label, seen[label], r))
        else:
            seen[label] = r
    if dup_items:
        log("发现重复商品行 %d 组：%s" % (
            len(dup_items), "；".join("%s（R%d 与 R%d）" % (n, a, b) for n, a, b in dup_items)))

    # ---------- 2. 逐人汇总 ----------
    orders = []
    for p in people:
        c = p["col"]
        raw_amount = ws.cell(meta_rows.get("应付金额", 0), c).value if "应付金额" in meta_rows else None
        amount = parse_amount(raw_amount)
        pay_status = ws.cell(meta_rows.get("支付状态", 0), c).value if "支付状态" in meta_rows else None
        order_type = ws.cell(meta_rows.get("订单类型", 0), c).value if "订单类型" in meta_rows else None
        raw_time = ws.cell(meta_rows.get("下单时间", 0), c).value if "下单时间" in meta_rows else None
        tm, tm_note = parse_time(raw_time)

        items, qty_total, issues = [], 0, []
        for r, label in item_rows:
            q = ws.cell(r, c).value
            if q is None or str(q).strip() in ("", "0"):
                continue
            qty = parse_amount(q)
            if qty is None:
                issues.append("商品「%s」数量无法识别：%r" % (label, q))
                continue
            if qty < 0:
                issues.append("商品「%s」数量为负：%s" % (label, qty))
            items.append("%s×%g" % (label, qty))
            qty_total += qty

        if not items and (amount or 0) > 0:
            issues.append("没有选任何商品，但应付金额为 %s" % amount)
        if not items and amount in (None, 0):
            issues.append("疑似空订单（无商品、无金额）")
        if amount is None:
            issues.append("应付金额无法识别：%r" % (raw_amount,))
        if tm is None:
            issues.append("下单时间%s" % (tm_note or "缺失"))
        elif tm_note:
            issues.append("下单时间%s" % tm_note)
        pay_str = str(pay_status).strip() if pay_status is not None else ""
        if pay_str not in ("已付款", "未付款", "已付", "未付", "部分付款", "已退款"):
            issues.append("支付状态不规范：%r" % (pay_status,))

        orders.append({
            "name": p["name"], "type": order_type, "amount": amount,
            "pay": pay_str or "缺失", "time": tm, "time_note": tm_note,
            "items": items, "qty": qty_total, "issues": issues,
        })

    paid = [o for o in orders if o["pay"] in ("已付款", "已付")]
    unpaid = [o for o in orders if o["pay"] in ("未付款", "未付", "部分付款", "缺失")]
    total_amount = sum(o["amount"] or 0 for o in orders)
    paid_amount = sum(o["amount"] or 0 for o in paid)
    unpaid_amount = sum(o["amount"] or 0 for o in unpaid)
    all_issues = sum(len(o["issues"]) for o in orders)

    log("对账完成：%d 人 ｜ 合计金额 %.2f ｜ 已付款 %d 人 %.2f ｜ 未付款 %d 人 %.2f ｜ 异常 %d 条" % (
        len(orders), total_amount, len(paid), paid_amount, len(unpaid), unpaid_amount, all_issues))

    # ---------- 3. 写出 orders_clean.xlsx ----------
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "订单明细"
    headers = ["姓名", "订单类型", "应付金额", "支付状态", "下单时间", "商品件数", "商品清单", "需跟进"]
    ws2.append(headers)
    for i, h in enumerate(headers, start=1):
        cell = ws2.cell(1, i)
        cell.font, cell.fill, cell.border = HEAD_FONT, HEAD_FILL, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for o in orders:
        row = [
            o["name"], o["type"], o["amount"], o["pay"],
            o["time"].strftime("%Y-%m-%d %H:%M") if o["time"] else (o["time_note"] or "缺失"),
            o["qty"], "；".join(o["items"]) or "（无）",
            "是" if o["issues"] else "",
        ]
        ws2.append(row)
        r = ws2.max_row
        for i in range(1, len(headers) + 1):
            ws2.cell(r, i).font = BODY_FONT
            ws2.cell(r, i).border = BORDER
        if o["issues"]:
            for i in range(1, len(headers) + 1):
                ws2.cell(r, i).fill = WARN_FILL
    widths = [12, 12, 12, 12, 20, 10, 60, 10]
    for i, w in enumerate(widths, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"
    p1 = os.path.join(out, "orders_clean.xlsx")
    wb2.save(p1)
    log("产出 %s" % p1)

    # ---------- 4. 写出 items_matrix.xlsx（转置后可读） ----------
    wb3 = openpyxl.Workbook()
    ws3 = wb3.active
    ws3.title = "商品矩阵"
    hdr = ["商品"] + [p["name"] for p in people] + ["合计件数"]
    ws3.append(hdr)
    for i in range(1, len(hdr) + 1):
        cell = ws3.cell(1, i)
        cell.font, cell.fill, cell.border = HEAD_FONT, HEAD_FILL, BORDER
        cell.alignment = Alignment(horizontal="center")
    for r, label in item_rows:
        row = [label]
        s = 0
        for p in people:
            q = parse_amount(ws.cell(r, p["col"]).value) or 0
            row.append(q if q else "")
            s += q
        row.append(s)
        ws3.append(row)
        for i in range(1, len(hdr) + 1):
            ws3.cell(ws3.max_row, i).font = BODY_FONT
            ws3.cell(ws3.max_row, i).border = BORDER
    tail = ["合计"] + [sum(parse_amount(ws.cell(r, p["col"]).value) or 0 for r, _ in item_rows) for p in people] + [""]
    ws3.append(tail)
    for i in range(1, len(hdr) + 1):
        c = ws3.cell(ws3.max_row, i)
        c.font = Font(name="微软雅黑", size=10, bold=True)
        c.border = BORDER
    ws3.column_dimensions["A"].width = 30
    p2 = os.path.join(out, "items_matrix.xlsx")
    wb3.save(p2)
    log("产出 %s" % p2)

    # ---------- 5. 写出 reconcile.md ----------
    md = []
    md.append("# 团购对账单\n")
    md.append("> 生成时间：%s ｜ 引擎：order_matrix v1.0" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    md.append("> 源文件：`%s` ｜ 全部数字由程序从原表计算，可在 `items_matrix.xlsx` 逐格核对。\n" % os.path.basename(args.input))
    md.append("## 一、总览\n")
    md.append("| 指标 | 数值 |")
    md.append("| --- | ---: |")
    md.append("| 下单人数 | %d 人 |" % len(orders))
    md.append("| 商品种类 | %d 种 |" % len(item_rows))
    md.append("| 合计件数 | %g 件 |" % sum(o["qty"] for o in orders))
    md.append("| 应付金额合计 | %.2f 元 |" % total_amount)
    md.append("| **已付款** | **%d 人 / %.2f 元** |" % (len(paid), paid_amount))
    md.append("| **未付款** | **%d 人 / %.2f 元** |" % (len(unpaid), unpaid_amount))
    md.append("| 数据异常 | %d 条（涉及 %d 人） |\n" % (all_issues, sum(1 for o in orders if o["issues"])))

    md.append("## 二、还没付款的人（要催的名单）\n")
    if unpaid:
        md.append("| 姓名 | 应付金额 | 状态 | 件数 |")
        md.append("| --- | ---: | --- | ---: |")
        for o in sorted(unpaid, key=lambda x: -(x["amount"] or 0)):
            md.append("| %s | %.2f | %s | %g |" % (o["name"], o["amount"] or 0, o["pay"], o["qty"]))
        md.append("")
    else:
        md.append("全部已付款。\n")

    md.append("## 三、每人明细\n")
    md.append("| 姓名 | 件数 | 应付 | 状态 | 商品清单 |")
    md.append("| --- | ---: | ---: | --- | --- |")
    for o in orders:
        md.append("| %s | %g | %s | %s | %s |" % (
            o["name"], o["qty"],
            "%.2f" % o["amount"] if o["amount"] is not None else "缺失",
            o["pay"], "；".join(o["items"]) or "（无）"))
    md.append("")

    md.append("## 四、数据问题清单\n")
    if all_issues == 0:
        md.append("未发现异常。\n")
    else:
        for o in orders:
            if o["issues"]:
                md.append("- **%s**：%s" % (o["name"], "；".join(o["issues"])))
        md.append("")
    if dup_items:
        md.append("### 原表结构问题\n")
        for n, a, b in dup_items:
            md.append("- 商品「%s」在原表出现了两次（第 %d 行与第 %d 行），统计会被重复计算。" % (n, a, b))
        md.append("")

    md.append("## 五、建议动作\n")
    md.append("1. 先催上表里的未付款名单，%d 人共 %.2f 元。" % (len(unpaid), unpaid_amount))
    md.append("2. 修原表的重复商品行，否则金额/件数会重复统计。")
    md.append("3. 支付状态建议统一为「已付款 / 未付款」两值，便于筛选。")
    md.append("")
    p3 = os.path.join(out, "reconcile.md")
    open(p3, "w", encoding="utf-8").write("\n".join(md))
    log("产出 %s" % p3)

    # ---------- 6. run.log ----------
    log("[完成] 总耗时 %.1f ms" % ((time.perf_counter() - t0) * 1000))
    open(os.path.join(out, "run.log"), "w", encoding="utf-8").write("\n".join(log_lines))

    print("\n汇总: %d 人 ｜ 合计 %.2f 元 ｜ 已付 %d 人 ｜ 未付 %d 人(%.2f 元) ｜ 异常 %d 条" % (
        len(orders), total_amount, len(paid), len(unpaid), unpaid_amount, all_issues))
    return 0


if __name__ == "__main__":
    sys.exit(main())