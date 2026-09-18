"""同文件夹 Excel 瓶类重量占比汇总工具。

将本程序放在待处理 Excel 文件所在文件夹，双击运行即可。
程序会在每个输入工作簿中定位“重量占比”列，按固定行号汇总并输出新文件。
"""

from __future__ import annotations

import os
import re
import sys
import traceback
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


OUTPUT_SUFFIX = "_瓶类重量占比汇总"
WEIGHT_HEADER = "重量占比"

# 第 8 条与第 12 条均为第 30 至 34 行，故仅保留一次。
SUMMARY_RULES: list[tuple[str, list[tuple[int, int]]]] = [
    ("白瓶", [(3, 11), (13, 15)]),
    ("3A白瓶", [(12, 12)]),
    ("小油壶", [(15, 15)]),
    ("绿瓶", [(16, 17)]),
    ("蓝瓶（18-19行）", [(18, 19)]),
    ("其他", [(20, 25), (41, 51)]),
    ("PE", [(27, 29)]),
    ("蓝瓶（30-34行）", [(30, 34)]),
    ("大油壶", [(35, 35)]),
    ("杂色PET", [(36, 37)]),
    ("阻隔瓶", [(38, 40)]),
    ("铝口瓶", [(52, 52)]),
]


class ProcessingError(Exception):
    """面向用户的输入或处理错误。"""


def app_directory() -> Path:
    """兼容直接运行脚本和 PyInstaller 打包后的 EXE。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def is_input_workbook(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
        and OUTPUT_SUFFIX not in path.stem
    )


def normalize_text(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def find_weight_column(workbook) -> tuple[str, int]:
    """在所有工作表中寻找“重量占比”列标题，返回首个匹配项。"""
    target = normalize_text(WEIGHT_HEADER)
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if target in normalize_text(cell.value):
                    return worksheet.title, cell.column
    raise ProcessingError("未找到“重量占比”列。请确认源 Excel 中包含该列标题。")


def to_number(value: object, cell_reference: str) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, bool):
        raise ProcessingError(f"{cell_reference} 的重量占比不是数值。")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    try:
        if text.endswith("%"):
            return float(text[:-1]) / 100
        return float(text)
    except ValueError as exc:
        raise ProcessingError(f"{cell_reference} 的重量占比“{value}”无法转换为数值。") from exc


def sum_ranges(worksheet, column: int, ranges: Iterable[tuple[int, int]]) -> float:
    total = 0.0
    for start_row, end_row in ranges:
        for row in range(start_row, end_row + 1):
            reference = f"{get_column_letter(column)}{row}"
            total += to_number(worksheet.cell(row=row, column=column).value, reference)
    return total


def create_summary_workbook(source_path: Path) -> tuple[Path, str, str]:
    """读取输入文件并创建汇总结果工作簿，不修改原文件。"""
    try:
        # data_only=True 用于读取 Excel 已保存的公式结果。
        source_book = load_workbook(source_path, data_only=True, read_only=False)
    except Exception as exc:
        raise ProcessingError(f"无法打开 {source_path.name}：{exc}") from exc

    sheet_name, weight_column = find_weight_column(source_book)
    source_sheet = source_book[sheet_name]
    if source_sheet.max_row < 52:
        raise ProcessingError(
            f"{source_path.name} 的“{sheet_name}”只有 {source_sheet.max_row} 行，"
            "不足以按第 52 行进行汇总。"
        )

    result_book = Workbook()
    result_sheet = result_book.active
    result_sheet.title = "汇总结果"
    result_sheet.append(["类别", "重量占比", "取数行号"])
    result_sheet.freeze_panes = "A2"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in result_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for category, ranges in SUMMARY_RULES:
        total = sum_ranges(source_sheet, weight_column, ranges)
        range_text = "、".join(
            str(start) if start == end else f"{start}-{end}" for start, end in ranges
        )
        result_sheet.append([category, total, range_text])

    total_row = result_sheet.max_row + 1
    result_sheet.cell(total_row, 1, "汇总项合计")
    result_sheet.cell(total_row, 2, f"=SUM(B2:B{total_row - 1})")
    result_sheet.cell(total_row, 1).font = Font(bold=True)
    result_sheet.cell(total_row, 2).font = Font(bold=True)

    for row in range(2, total_row + 1):
        result_sheet.cell(row, 2).number_format = "0.00%"
    result_sheet.column_dimensions["A"].width = 22
    result_sheet.column_dimensions["B"].width = 16
    result_sheet.column_dimensions["C"].width = 18

    notes = result_book.create_sheet("处理说明")
    notes.append(["项目", "内容"])
    notes.append(["输入文件", source_path.name])
    notes.append(["来源工作表", sheet_name])
    notes.append(["重量占比列", get_column_letter(weight_column)])
    notes.append(["固定规则", "第 8 条与第 12 条重复，仅输出一次第 30-34 行蓝瓶结果。"])
    notes.column_dimensions["A"].width = 18
    notes.column_dimensions["B"].width = 70
    for cell in notes[1]:
        cell.fill = header_fill
        cell.font = header_font
    for row in notes.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    output_path = source_path.with_name(f"{source_path.stem}{OUTPUT_SUFFIX}.xlsx")
    try:
        result_book.save(output_path)
    except PermissionError as exc:
        raise ProcessingError(
            f"无法保存 {output_path.name}。请先关闭已打开的同名结果文件后重试。"
        ) from exc
    return output_path, sheet_name, get_column_letter(weight_column)


def show_message(title: str, message: str, *, error: bool = False) -> None:
    """打包为无控制台 EXE 后仍能显示清晰结果。"""
    try:
        import tkinter.messagebox as messagebox

        root = __import__("tkinter").Tk()
        root.withdraw()
        if error:
            messagebox.showerror(title, message, parent=root)
        else:
            messagebox.showinfo(title, message, parent=root)
        root.destroy()
    except Exception:
        print(f"{title}\n{message}")


def main() -> int:
    folder = app_directory()
    inputs = sorted(path for path in folder.iterdir() if is_input_workbook(path))
    if not inputs:
        show_message(
            "未找到输入文件",
            "请将本程序与需要处理的 .xlsx 或 .xlsm 文件放在同一文件夹，然后双击运行。",
            error=True,
        )
        return 1

    successes: list[str] = []
    failures: list[str] = []
    for input_path in inputs:
        try:
            output_path, sheet_name, column = create_summary_workbook(input_path)
            successes.append(
                f"{input_path.name} → {output_path.name}\n"
                f"  取数：{sheet_name} 工作表，{column} 列“重量占比”"
            )
        except ProcessingError as exc:
            failures.append(f"{input_path.name}：{exc}")

    message = "\n\n".join(successes)
    if failures:
        message += ("\n\n未处理文件：\n" + "\n".join(failures))
    show_message(
        "重量占比汇总完成" if successes else "重量占比汇总未完成",
        message,
        error=not bool(successes),
    )
    return 0 if successes else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # 防止双击程序无提示退出。
        traceback.print_exc()
        show_message("程序发生错误", f"发生未预期错误：{exc}", error=True)
        raise SystemExit(1)
