"""GUI 界面 — PyQt5 桌面应用"""
import sys
import os
import asyncio
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Set

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QDateEdit, QFileDialog, QTextEdit,
    QProgressBar, QMessageBox, QGroupBox, QFrame, QTableWidget,
    QTableWidgetItem, QTabWidget, QSplitter, QStatusBar
)
from PyQt5.QtCore import QDate, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QColor

import yaml

from utils.logger import setup_logger, GuiLogHandler
from utils.fuzzy_match import normalize_text


class WorkerThread(QThread):
    """后台处理线程"""
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)  # current, total
    finished_signal = pyqtSignal(str, str)  # missing_file, filtered_file
    error_signal = pyqtSignal(str)

    def __init__(self, config, date_from, date_to, agent_email_path=None):
        super().__init__()
        self.config = config
        self.date_from = date_from
        self.date_to = date_to
        self.agent_email_path = agent_email_path
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            import logging
            logger = logging.getLogger("mail_audit")
            logger.setLevel(logging.DEBUG)
            handler = GuiLogHandler(self._log)
            handler.setLevel(logging.INFO)
            fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
            handler.setFormatter(fmt)
            logger.addHandler(handler)

            logger.info("=" * 50)
            logger.info("开始运行漏单审核流程")
            logger.info("=" * 50)

            # 加载参照表
            from openpyxl import load_workbook
            ref_cfg = self.config["reference_tables"]

            # 加载附件二: 内部邮箱
            internal_emails = self._load_internal_emails(ref_cfg["internal_emails"])
            logger.info(f"内部邮箱表已加载: {len(internal_emails)} 条")

            # 加载附件三: 代理邮箱
            agent_email_path = self.agent_email_path or ref_cfg["agent_emails"]
            agent_map = self._load_agent_emails(agent_email_path)
            logger.info(f"代理邮箱表已加载: {len(agent_map)} 条")

            # 加载附件四: 项目名称
            project_table = self._load_project_names(ref_cfg["project_names"])
            logger.info(f"项目名称表已加载: {len(project_table)} 条")

            if self._stop:
                return

            # M1: 邮件读取
            from modules.mail_reader import MailReader
            reader = MailReader(self.config["email"], logger)
            self_email = self.config["email"]["address"]
            mails = reader.fetch_mails(
                self.date_from, self.date_to, self_email,
                progress_callback=lambda c, t: self.progress_signal.emit(c, t)
            )
            logger.info(f"邮件拉取完成: 共 {len(mails)} 封")

            if self._stop:
                return

            # M2: 邮件过滤
            from modules.mail_filter import MailFilter
            filt = MailFilter(internal_emails, logger)
            valid_mails, filtered_mails = filt.filter_mails(mails)
            logger.info(f"过滤完成: 有效 {len(valid_mails)} 封, 过滤 {len(filtered_mails)} 封")

            if self._stop:
                return

            # M3: 字段提取
            from modules.field_extractor import FieldExtractor
            extractor = FieldExtractor(agent_map, project_table, logger)
            all_rows = []
            for idx, mail in enumerate(valid_mails, 1):
                if self._stop:
                    return
                rows = extractor.extract_fields(mail)
                all_rows.extend(rows)
                logger.info(f"字段提取: 第{idx}/{len(valid_mails)}封 — "
                           f"客户='{rows[0].get('客户','')}', 代理='{rows[0].get('代理','')}'")
                self.progress_signal.emit(idx, len(valid_mails))

            if self._stop:
                return

            # M4: 项目标准化拆分 (M3中已完成基础提取, 这里做组合拆分)
            from modules.project_normalizer import ProjectNormalizer
            normalizer = ProjectNormalizer(project_table, logger)
            normalized_rows = []
            for row in all_rows:
                proj_raw = row.get("项目", "")
                if proj_raw and "/" in proj_raw or "+" in proj_raw:
                    split_results = normalizer.normalize_and_split(
                        proj_raw, row.get("subject", ""), row.get("body_text", "")
                    )
                    if len(split_results) > 1:
                        for sr in split_results:
                            new_row = row.copy()
                            new_row["项目"] = sr["standard_name"]
                            new_row["项目原始值"] = proj_raw
                            normalized_rows.append(new_row)
                        continue
                normalized_rows.append(row)

            logger.info(f"项目标准化完成: {len(normalized_rows)} 行")
            all_rows = normalized_rows

            if self._stop:
                return

            # M5: 工单系统比对
            from modules.workorder_checker import WorkOrderChecker
            checker = WorkOrderChecker(self.config["workorder"], logger)

            # 运行 async 登录
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            login_ok = loop.run_until_complete(checker.login())
            if not login_ok:
                logger.error("工单系统登录失败，跳过比对，直接输出未比对结果")
                for row in all_rows:
                    row["是否已录单"] = "未比对"
                    row["匹配状态"] = "工单系统未连接"
                    row["查询时间戳"] = datetime.now()
            else:
                # 收集所有唯一客户名，批量查询
                unique_customers = list(set(r.get("客户", "") for r in all_rows if r.get("客户")))
                all_workorders = []
                for idx, cust in enumerate(unique_customers, 1):
                    if self._stop:
                        return
                    logger.info(f"工单查询: {idx}/{len(unique_customers)} — {cust}")
                    orders = loop.run_until_complete(checker.search_workorders(cust))
                    all_workorders.extend(orders)
                    self.progress_signal.emit(idx, len(unique_customers))

                # 比对
                all_rows = checker.match_records(all_rows, all_workorders)
                logger.info(f"工单比对完成: 漏单 {sum(1 for r in all_rows if r.get('是否已录单')=='否')} 条")
                loop.run_until_complete(checker.close())

            loop.close()

            if self._stop:
                return

            # M6: 输出 Excel
            from modules.excel_writer import ExcelWriter
            writer = ExcelWriter(self.config["output"]["dir"], logger)
            missing_file = writer.write_missing_report(all_rows)
            filtered_file = writer.write_filtered_report(filtered_mails)

            logger.info(f"输出完成!")
            logger.info(f"  漏单清单: {missing_file}")
            logger.info(f"  过滤清单: {filtered_file}")

            self.finished_signal.emit(missing_file, filtered_file)

        except Exception as e:
            import traceback
            self.error_signal.emit(f"{e}\n{traceback.format_exc()}")

    def _log(self, msg):
        self.log_signal.emit(msg)

    def _load_internal_emails(self, path) -> Set[str]:
        """加载附件二: 内部邮箱表"""
        if not os.path.exists(path):
            return set()
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        emails = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            for cell in row:
                if cell and "@" in str(cell):
                    emails.add(str(cell).strip().lower())
        wb.close()
        return emails

    def _load_agent_emails(self, path) -> Dict[str, dict]:
        """加载附件三: 代理邮箱表"""
        if not os.path.exists(path):
            return {}
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        result = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            agent = str(row[0]).strip() if row[0] else ""
            short = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            email = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            if email and "@" in email:
                # 一个代理可能有多个邮箱(分号分隔)
                for em in email.replace("；", ";").split(";"):
                    em = em.strip().lower()
                    if em:
                        result[em] = {"代理": agent, "代理简称": short, "收件人邮箱": email}
        wb.close()
        return result

    def _load_project_names(self, path) -> List[dict]:
        """加载附件四: 项目名称表"""
        if not os.path.exists(path):
            return []
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        result = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            entry = {}
            keys = ["项目编号", "项目名称", "国家", "业务类型"]
            for idx, key in enumerate(keys):
                if idx < len(row) and row[idx]:
                    entry[key] = str(row[idx]).strip()
            if entry.get("项目名称"):
                result.append(entry)
        wb.close()
        return result


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = None
        self.worker = None
        self.init_ui()
        self.load_config()

    def init_ui(self):
        self.setWindowTitle("邮件&系统漏单审核自动化工具 v1.0")
        self.setMinimumSize(900, 650)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        # === 顶部控制栏 ===
        ctrl_group = QGroupBox("控制面板")
        ctrl_layout = QHBoxLayout(ctrl_group)

        ctrl_layout.addWidget(QLabel("时间范围:"))
        self.date_from = QDateEdit()
        self.date_from.setDate(QDate.currentDate().addDays(-30))
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.setCalendarPopup(True)
        ctrl_layout.addWidget(self.date_from)

        ctrl_layout.addWidget(QLabel("至"))
        self.date_to = QDateEdit()
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setDisplayFormat("yyyy-MM-dd")
        self.date_to.setCalendarPopup(True)
        ctrl_layout.addWidget(self.date_to)

        ctrl_layout.addSpacing(20)

        self.btn_import_agent = QPushButton("邮箱汇总表导入")
        self.btn_import_agent.clicked.connect(self.import_agent_table)
        ctrl_layout.addWidget(self.btn_import_agent)
        self.agent_table_label = QLabel("未导入")
        self.agent_table_label.setStyleSheet("color: #999;")
        ctrl_layout.addWidget(self.agent_table_label)

        ctrl_layout.addStretch()

        self.btn_run = QPushButton("开始运行")
        self.btn_run.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; font-size: 14px; font-weight: bold; padding: 6px 20px; }"
            "QPushButton:hover { background-color: #45a049; }"
        )
        self.btn_run.clicked.connect(self.start_run)
        ctrl_layout.addWidget(self.btn_run)

        self.btn_stop = QPushButton("停止")
        self.btn_stop.setStyleSheet(
            "QPushButton { background-color: #f44336; color: white; font-size: 14px; padding: 6px 20px; }"
            "QPushButton:hover { background-color: #d32f2f; }"
        )
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_run)
        ctrl_layout.addWidget(self.btn_stop)

        layout.addWidget(ctrl_group)

        # === 进度条 ===
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        # === 日志面板 ===
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet(
            "QTextEdit { background-color: #1e1e1e; color: #d4d4d4; }"
        )
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group, stretch=1)

        # === 底部按钮 ===
        bottom_widget = QWidget()
        bottom_layout = QHBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_open_missing = QPushButton("查看漏单清单")
        self.btn_open_missing.setEnabled(False)
        self.btn_open_missing.clicked.connect(lambda: self.open_excel(self.missing_file))
        bottom_layout.addWidget(self.btn_open_missing)

        self.btn_open_filtered = QPushButton("查看过滤清单")
        self.btn_open_filtered.setEnabled(False)
        self.btn_open_filtered.clicked.connect(lambda: self.open_excel(self.filtered_file))
        bottom_layout.addWidget(self.btn_open_filtered)

        bottom_layout.addStretch()
        btn_exit = QPushButton("退出")
        btn_exit.clicked.connect(self.close)
        bottom_layout.addWidget(btn_exit)
        layout.addWidget(bottom_widget)

        self.missing_file = ""
        self.filtered_file = ""

    def load_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                self.config = yaml.safe_load(f)
        else:
            QMessageBox.critical(self, "错误", "找不到 config.yaml 配置文件!")
            self.config = {}

    def import_agent_table(self):
        """导入邮箱汇总表"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择邮箱汇总表 Excel 文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.agent_email_path = path
            fname = os.path.basename(path)
            self.agent_table_label.setText(f"已导入: {fname}")
            self.agent_table_label.setStyleSheet("color: #4CAF50;")
            self.log(f"邮箱汇总表已导入: {fname}")

    def start_run(self):
        if not self.config:
            QMessageBox.critical(self, "错误", "配置文件未加载!")
            return

        date_from = datetime(
            self.date_from.date().year(), self.date_from.date().month(), self.date_from.date().day()
        )
        date_to = datetime(
            self.date_to.date().year(), self.date_to.date().month(), self.date_to.date().day()
        ) + timedelta(days=1)

        agent_path = getattr(self, "agent_email_path", None)

        self.worker = WorkerThread(self.config, date_from, date_to, agent_path)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(
            lambda c, t: self.progress.setValue(int(c * 100 / t) if t else 0)
        )
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.start()

        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_open_missing.setEnabled(False)
        self.btn_open_filtered.setEnabled(False)
        self.progress.setValue(0)

    def stop_run(self):
        if self.worker:
            self.worker.stop()
            self.log("用户请求停止运行...")
            self.btn_stop.setEnabled(False)

    def on_finished(self, missing_file, filtered_file):
        self.missing_file = missing_file
        self.filtered_file = filtered_file
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_open_missing.setEnabled(True)
        self.btn_open_filtered.setEnabled(True)
        self.progress.setValue(100)
        self.log("=" * 50)
        self.log("运行完成!")
        self.log(f"  漏单清单: {missing_file}")
        self.log(f"  过滤清单: {filtered_file}")

    def on_error(self, error_msg):
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.log(f"[ERROR] {error_msg}")
        QMessageBox.critical(self, "运行错误", error_msg)

    def log(self, msg):
        self.log_text.append(msg)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def open_excel(self, filepath):
        if filepath and os.path.exists(filepath):
            os.startfile(filepath)
        else:
            QMessageBox.warning(self, "提示", "文件不存在")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
