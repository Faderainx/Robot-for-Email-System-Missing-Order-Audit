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

    # --------------- 文件名规范：稳定名 + 时序副本 ---------------

    @staticmethod
    def _stable_and_timestamp(stable_name: str, ext: str = ".xlsx") -> tuple:
        """返回 (稳定名路径, 时序副本路径). 副本不重复时与稳定名相同写入。"""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        # 时序副本: {stable}_{ts}.xlsx, 与稳定名同目录
        base, suf = os.path.splitext(stable_name)
        ts_path = f"{base}_{ts}{suf}"
        return stable_name, ts_path

    def write_stage1_outputs(self, all_rows: List[Dict], filtered_mails: List[Dict]) -> tuple:
        """
        阶段一输出三份:
          1. to_workorder_list.xlsx    (稳定名, 时序副本) — 有把握查工单的:代理命中1条 / 客户非空
          2. to_review_list.xlsx       (稳定名, 时序副本) — 需人工补全:代理未命中/候选≥2
          3. filtered_mail_record.xlsx  (稳定名, 时序副本) — 被过滤邮件审计
        返回 (to_workorder_path, to_review_path)
        """
        to_work, to_review = self._partition_rows(all_rows)

        # 1) to_workorder_list.xlsx
        wo_stable, _ = self._write_to_workorder_list(to_work)
        # 2) to_review_list.xlsx
        rv_stable, _ = self._write_to_review_list(to_review)
        # 3) filtered_mail_record.xlsx (稳定名 + 时序副本)
        fr_stable, fr_ts = self._stable_and_timestamp(
            os.path.join(self.output_dir, "filtered_mail_record.xlsx")
        )
        self._write_filtered_mail_record(filtered_mails, fr_stable, fr_ts)

        self._log(f"[stage1] 待查清单: {wo_stable}")
        self._log(f"[stage1] 漏单复查: {rv_stable}")
        self._log(f"[stage1] 过滤日志: {fr_stable}")
        return wo_stable, rv_stable

    def write_workorder_check_result(
        self, all_rows: List[Dict], to_query: List[Dict], skipped: List[Dict] = None
    ) -> tuple:
        """
        阶段二输出 workorder_check_result.xlsx — 包含原始邮件信息 + 工单查询结果 + 跳过条目标注
        返回 (stable_path, ts_path)
        """
        stable, ts = self._stable_and_timestamp(
            os.path.join(self.output_dir, "workorder_check_result.xlsx")
        )
        self._do_write_workorder_check(all_rows, to_query, skipped, stable)
        # 时序副本
        try:
            import shutil
            shutil.copy(stable, ts)
            self._log(f"[stage2] 时序副本: {ts}")
        except Exception as e:
            self._log(f"[stage2] 时序副本复制失败: {e}", "warning")
        self._log(f"[stage2] 工单核对结果: {stable}")
        return stable, ts

    def write_all_outputs(self, all_rows: List[Dict], filtered_mails: List[Dict]) -> tuple:
        """一站式输出: 沿用原 v1.0 行为 — 漏单清单 + 过滤清单 (兼容旧调用者)"""
        missing = self.write_missing_report(all_rows)
        filtered = self.write_filtered_report(filtered_mails)
        return missing, filtered

    # --------------- 私有: 行分区与各文件写入 ---------------

    @staticmethod
    def _partition_rows(rows: List[Dict]) -> tuple:
        """
        按代理匹配结果拆:
          to_work: 代理命中1条 (或代理已确认) — 进入待查清单
          to_review: 代理未命中 OR 候选≥2 + 置信度需人工 — 漏单复查
        """
        to_work, to_review = [], []
        for r in rows:
            agent = (r.get("代理") or "").strip()
            confidence = (r.get("置信度") or "").strip()
            match_way = (r.get("代理匹配方式") or "").strip()
            # 命中0条 或 候选≥2 (多匹配) — 进复查
            if not agent or "多匹配" in match_way or confidence == "需人工确认":
                to_review.append(r)
            else:
                to_work.append(r)
        return to_work, to_review

    def _write_to_workorder_list(self, rows: List[Dict]) -> tuple:
        """to_workorder_list.xlsx — 阶段二唯一输入。稳定名 + 时序副本。"""
        stable, ts = self._stable_and_timestamp(
            os.path.join(self.output_dir, "to_workorder_list.xlsx")
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "工单待查"

        headers = [
            "发件人邮箱", "发件日期", "邮件主题",
            "代理", "代理匹配方式",
            "客户公司名称", "客户提取来源",
            "标准化项目名称", "项目原始值",
            "需求", "置信度", "数据来源",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        for row_idx, row in enumerate(rows, 2):
            values = [
                row.get("sender_email", ""),
                self._format_date(row.get("date")),
                row.get("subject", ""),
                row.get("代理", ""),
                row.get("代理匹配方式", ""),
                row.get("客户", ""),
                row.get("客户提取来源", ""),
                row.get("项目", ""),
                row.get("项目原始值", ""),
                row.get("需求", ""),
                row.get("置信度", ""),
                row.get("数据来源", row.get("客户提取来源", "")),  # 兼容字段
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            # 颜色标记: confidence=low -> 蓝; 多匹配 -> 橙
            if row.get("置信度") == "low" or "待确认" in (row.get("代理匹配方式") or ""):
                for col in [4, 6, 8]:
                    ws.cell(row=row_idx, column=col).fill = COLOR_BLUE_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_BLUE_FONT

        col_widths = [25, 20, 40, 18, 18, 28, 18, 22, 22, 10, 12, 22]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[chr(64 + idx)].width = width
        ws.freeze_panes = "A2"
        wb.save(stable)

        # 时序副本
        try:
            import shutil
            shutil.copy(stable, ts)
        except Exception as e:
            self._log(f"待查清单时序副本复制失败: {e}", "warning")
        self._log(f"工单待查清单: {stable} ({len(rows)} 行)")
        return stable, ts

    def _write_to_review_list(self, rows: List[Dict]) -> tuple:
        """to_review_list.xlsx — 漏单复查: 代理未命中/候选≥2 等需人工补全。稳定名 + 时序副本。"""
        stable, ts = self._stable_and_timestamp(
            os.path.join(self.output_dir, "to_review_list.xlsx")
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "漏单复查"

        headers = [
            "发件人邮箱", "发件日期", "邮件主题",
            "代理(可能空)", "候选代理(匹配度降序)", "代理匹配方式",
            "客户公司名称", "标准化项目名称", "需求",
            "置信度", "数据来源", "待人工补全",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        for row_idx, row in enumerate(rows, 2):
            # 候选代理: 多匹配时写候选列表；空时写 "未匹配"
            agent_candidates = row.get("候选代理", "")
            if not agent_candidates:
                candidates_str = ""
            else:
                if isinstance(agent_candidates, list):
                    candidates_str = " | ".join(
                        f"{c.get('代理','')}({c.get('score','')})"
                        if isinstance(c, dict) else str(c)
                        for c in agent_candidates
                    )
                else:
                    candidates_str = str(agent_candidates)

            values = [
                row.get("sender_email", ""),
                self._format_date(row.get("date")),
                row.get("subject", ""),
                row.get("代理", ""),
                candidates_str,
                row.get("代理匹配方式", ""),
                row.get("客户", ""),
                row.get("项目", ""),
                row.get("需求", ""),
                row.get("置信度", ""),
                row.get("数据来源", row.get("客户提取来源", "")),
                "请人工补全代理后导入 to_workorder_list.xlsx",
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            # 整行橙色 — 需人工复查
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = COLOR_ORANGE_FILL
                ws.cell(row=row_idx, column=col).font = COLOR_ORANGE_FONT

        col_widths = [25, 20, 40, 18, 35, 22, 28, 22, 10, 12, 22, 30]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[chr(64 + idx)].width = width
        ws.freeze_panes = "A2"
        wb.save(stable)

        try:
            import shutil
            shutil.copy(stable, ts)
        except Exception as e:
            self._log(f"漏单复查时序副本复制失败: {e}", "warning")
        self._log(f"漏单复查清单: {stable} ({len(rows)} 行)")
        return stable, ts

    def _write_filtered_mail_record(self, filtered_mails: List[Dict],
                                     stable_path: str, ts_path: str) -> str:
        """filtered_mail_record.xlsx — 过滤日志审计。稳定名 + 时序副本。"""
        wb = Workbook()
        ws = wb.active
        ws.title = "过滤日志"

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
        wb.save(stable_path)

        try:
            import shutil
            shutil.copy(stable_path, ts_path)
        except Exception as e:
            self._log(f"过滤日志时序副本复制失败: {e}", "warning")
        self._log(f"过滤日志: {stable_path} ({len(filtered_mails)} 行)")
        return stable_path

    def _do_write_workorder_check(self, all_rows, to_query, skipped, output_path):
        """workorder_check_result.xlsx — 阶段二完整输出含跳过说明"""
        wb = Workbook()
        ws = wb.active
        ws.title = "工单核对"

        headers = [
            "发件人邮箱", "发件日期", "邮件主题", "正文(精简)",
            "代理", "客户", "标准化项目", "需求", "置信度",
            "是否已录单", "工单日期", "匹配状态",
            "RPA查询状态", "跳过后说明", "查询时间戳",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = COLOR_HEADER_FILL
            cell.font = COLOR_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

        # 已查: 写入to_query的数据 + 查询结果
        queried = {id(r): r for r in to_query}
        skipped_by_id = {id(s): s for s in (skipped or [])}

        for row_idx, row in enumerate(all_rows, 2):
            if id(row) in queried:
                q = queried[id(row)]
                values = [
                    q.get("sender_email", ""),
                    self._format_date(q.get("date")),
                    q.get("subject", ""),
                    (q.get("body_text", "") or "")[:200],
                    q.get("代理", ""),
                    q.get("客户", ""),
                    q.get("项目", ""),
                    q.get("需求", ""),
                    q.get("置信度", ""),
                    q.get("是否已录单", ""),
                    self._format_date(q.get("工单日期")),
                    q.get("匹配状态", ""),
                    "已查询",
                    "",
                    self._format_date(q.get("查询时间戳")),
                ]
            elif id(row) in skipped_by_id:
                s = skipped_by_id[id(row)]
                values = [
                    s.get("sender_email", ""),
                    self._format_date(s.get("date")),
                    s.get("subject", ""),
                    (s.get("body_text", "") or "")[:200],
                    s.get("代理", ""),
                    s.get("客户", ""),
                    s.get("项目", ""),
                    s.get("需求", ""),
                    s.get("置信度", ""),
                    "未查询",
                    "",
                    "前置校验跳过",
                    "跳过",
                    s.get("_skip_reason", ""),
                    "",
                ]
            else:
                # 不应该出现, 但兜底
                values = [""] * len(headers)
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)

            # 颜色: 已查询但漏单 -> 红; 前置跳过 -> 灰
            is_missing = values[9] == "否"
            is_skipped = values[12] == "跳过"
            if is_missing:
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col).fill = COLOR_RED_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_RED_FONT
            elif is_skipped:
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col).fill = COLOR_GRAY_FILL
                    ws.cell(row=row_idx, column=col).font = COLOR_GRAY_FONT

        col_widths = [25, 20, 40, 50, 18, 28, 22, 10, 12, 12, 20, 18, 14, 30, 20]
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[chr(64 + idx) if idx <= 26 else "A"].width = width
        ws.freeze_panes = "A2"
        wb.save(output_path)

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
