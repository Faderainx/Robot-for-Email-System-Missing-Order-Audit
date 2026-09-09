"""M6 Excel 输出模块 — 生成双 Excel 表格（漏单清单 + 过滤清单）"""
import os
from datetime import datetime
from typing import List, Dict

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# 颜色定义
COLOR_RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
COLOR_RED_FONT = Font(color="9C0006")
COLOR_YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
COLOR_YELLOW_FONT = Font(color="9C6500")
COLOR_BLUE_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
COLOR_BLUE_FONT = Font(color="1F4E79")
COLOR_ORANGE_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
COLOR_ORANGE_FONT = Font(color="833C00")
COLOR_GRAY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
COLOR_GRAY_FONT = Font(color="595959")
COLOR_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
COLOR_HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)

THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


class ExcelWriter:
    def __init__(self, output_dir: str = "output", logger=None):
        self.output_dir = output_dir
        self.logger = logger
        os.makedirs(output_dir, exist_ok=True)

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def write_missing_report(self, rows: List[Dict]) -> str:
        """生成漏单清单 Excel"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.output_dir, f"漏单清单_{timestamp}.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "漏单清单"

        headers = [
            "发件人邮箱", "发件日期", "主题", "正文(精简)",
            "代理", "代理匹配方式", "客户", "客户提取来源",
            "项目", "项目原始值", "需求",
            "是否已录单", "工单日期", "匹配状态",
            "查询时间戳", "备注",
        ]

        # 写表头
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        # 写数据
        for row_idx, row in enumerate(rows, 2):
            values = [
                row.get("sender_email", ""),
                self._format_date(row.get("date")),
                row.get("subject", ""),
                row.get("body_text", "")[:200],
                row.get("代理", ""),
                row.get("代理匹配方式", ""),
                row.get("客户", ""),
                row.get("客户提取来源", ""),
                row.get("项目", ""),
                row.get("项目原始值", ""),
                row.get("需求", ""),
                row.get("是否已录单", ""),
                self._format_date(row.get("工单日期")),
                row.get("匹配状态", ""),
                self._format_date(row.get("查询时间戳")),
                "",
            ]

            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)

            # 颜色标记
            self._apply_row_color(ws, row_idx, len(headers), row)

        # 调整列宽
        col_widths = [25, 20, 40, 50, 15, 18, 25, 18, 18, 25, 10, 12, 20, 18, 20, 15]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[chr(64 + idx) if idx <= 26 else "A"].width = width

        ws.freeze_panes = "A2"

        wb.save(filename)
        self._log(f"漏单清单已生成: {filename}")
        return filename

    def write_filtered_report(self, filtered_mails: List[Dict]) -> str:
        """生成过滤清单 Excel"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.output_dir, f"过滤清单_{timestamp}.xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "过滤清单"

        headers = ["发件人邮箱", "发件日期", "主题", "过滤原因", "处理时间戳"]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for row_idx, mail in enumerate(filtered_mails, 2):
            values = [
                mail.get("sender_email", ""),
                self._format_date(mail.get("date")),
                mail.get("subject", ""),
                mail.get("filter_reason", ""),
                now_str,
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        col_widths = [25, 20, 50, 20, 20]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[chr(64 + idx)].width = width

        ws.freeze_panes = "A2"
        wb.save(filename)
        self._log(f"过滤清单已生成: {filename}")
        return filename

    def _apply_row_color(self, ws, row_idx: int, col_count: int, row: Dict):
        """按规则给行上色"""
        status = row.get("匹配状态", "")
        is_missing = row.get("是否已录单") == "否"
        confidence = row.get("置信度", "")
        filter_status = row.get("filter_status", "")
        has_date_anomaly = any(
            wo.get("日期异常") for wo in row.get("工单记录", [])
        )

        if filter_status == "uncertain":
            for col in range(1, col_count + 1):
                ws.cell(row=row_idx, column=col).fill = COLOR_GRAY_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_GRAY_FONT
            return

        if is_missing:
            for col in range(1, col_count + 1):
                ws.cell(row=row_idx, column=col).fill = COLOR_RED_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_RED_FONT
            return

        if "多匹配" in status:
            cell = ws.cell(row=row_idx, column=14)
            cell.fill = COLOR_ORANGE_FILL
            cell.font = COLOR_ORANGE_FONT
            return

        if has_date_anomaly:
            cell = ws.cell(row=row_idx, column=13)
            cell.fill = COLOR_YELLOW_FILL
            cell.font = COLOR_YELLOW_FONT
            return

        if confidence == "low" or "待确认" in row.get("代理匹配方式", ""):
            for col in [5, 7, 9]:
                ws.cell(row=row_idx, column=col).fill = COLOR_BLUE_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_BLUE_FONT

    def _format_date(self, dt) -> str:
        if dt is None:
            return ""
        if isinstance(dt, str):
            return dt
        if isinstance(dt, datetime):
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return str(dt)
