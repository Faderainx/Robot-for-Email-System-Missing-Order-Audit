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
    QTableWidgetItem, QTabWidget, QSplitter, QStatusBar,
    QLineEdit, QFormLayout, QRadioButton, QButtonGroup
)
from PyQt5.QtCore import QDate, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QColor

import yaml
import json

from utils.logger import setup_logger, GuiLogHandler
from utils.fuzzy_match import normalize_text
from modules.llm_intent import LLMIntentClient

SESSION_FILE = os.path.join(os.path.dirname(__file__), "session_state.json")


class WorkerThread(QThread):
    """后台处理线程 — 支持独立运行 stage1(邮件解析) / stage2(工单查询) / all(全流程)"""
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)  # current, total
    finished_signal = pyqtSignal(str, str)  # primary_file, secondary_file
    error_signal = pyqtSignal(str)

    def __init__(self, config, mode="all", date_from=None, date_to=None,
                 agent_email_path=None, internal_email_path=None, project_table_path=None,
                 stage2_input_path=None):
        super().__init__()
        self.config = config
        self.mode = mode  # "stage1" | "stage2" | "all"
        self.date_from = date_from
        self.date_to = date_to
        self.agent_email_path = agent_email_path
        self.internal_email_path = internal_email_path
        self.project_table_path = project_table_path
        self.stage2_input_path = stage2_input_path  # 阶段二输入：阶段一输出的 xlsx
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            if self.mode == "stage2":
                self._run_stage2()
            elif self.mode == "stage1":
                self._run_stage1()
            else:  # "all"
                self._run_all()
        except Exception as e:
            import traceback
            self.error_signal.emit(f"{e}\n{traceback.format_exc()}")

    def _run_stage1(self):
        """阶段一: 邮件拉取→过滤→LLM意图分类→字段提取→OCR懒加载兜底→项目标准化→输出三份stage1产物"""
        import logging
        logger = logging.getLogger("mail_audit")
        logger.setLevel(logging.DEBUG)
        handler = GuiLogHandler(self._log)
        handler.setLevel(logging.INFO)
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
        handler.setFormatter(fmt)
        logger.addHandler(handler)

        logger.info("=" * 50)
        logger.info("[阶段一] 邮件过滤解析 (不含工单查询)")
        logger.info("=" * 50)

        all_rows, filtered_mails = self._run_pipeline_stage1(logger)
        if all_rows is None:
            return  # 用户停止

        if self._stop:
            return

        # 输出三份产物
        from modules.excel_writer import ExcelWriter
        writer = ExcelWriter(self.config["output"]["dir"], logger)
        primary, secondary = writer.write_stage1_outputs(all_rows, filtered_mails)
        logger.info("=" * 50)
        logger.info(f"[阶段一完成] 待查清单: {primary}")
        logger.info(f"[阶段一完成] 漏单复查: {secondary}")

        self.finished_signal.emit(primary, secondary)

    def _run_stage2(self):
        """阶段二: 读取阶段一输出 → 前置校验 → 工单查询 → 输出 workorder_check_result.xlsx"""
        import logging
        logger = logging.getLogger("mail_audit")
        logger.setLevel(logging.DEBUG)
        handler = GuiLogHandler(self._log)
        handler.setLevel(logging.INFO)
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
        handler.setFormatter(fmt)
        logger.addHandler(handler)

        logger.info("=" * 50)
        logger.info("[阶段二] 工单系统比对 (读阶段一输出)")
        logger.info("=" * 50)

        if not self.stage2_input_path or not os.path.exists(self.stage2_input_path):
            logger.error(f"阶段二输入文件不存在: {self.stage2_input_path}")
            self.error_signal.emit(f"阶段二输入文件不存在: {self.stage2_input_path}")
            return

        from openpyxl import load_workbook
        wb = load_workbook(self.stage2_input_path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            logger.error("阶段二输入文件为空")
            return
        headers = [str(c) if c is not None else "" for c in rows[0]]
        all_rows = [dict(zip(headers, [str(c) if c is not None else "" for c in r])) for r in rows[1:]]
        logger.info(f"已读取 {len(all_rows)} 行工单待查清单")

        if self._stop:
            return

        # 前置校验: 跳过代理空+置信度需人工确认; (公司,项目) 去重
        from modules.workorder_checker import WorkOrderChecker
        to_query, skipped = WorkOrderChecker.preprocess_rows(all_rows)
        if skipped:
            logger.warning(f"前置校验跳过 {len(skipped)} 条 (代理空+需人工确认 / 重复):")
            for s in skipped[:5]:
                logger.warning(f"  - 跳过: {s.get('客户','')} / {s.get('项目','')} — {s.get('_skip_reason','')}")
        logger.info(f"待RPA查询: {len(to_query)} 条 (原始 {len(all_rows)} 条)")

        if not to_query:
            logger.info("无可查询条目，跳过RPA")
            # 即使全跳空也输出空结果文件
            from modules.excel_writer import ExcelWriter
            writer = ExcelWriter(self.config["output"]["dir"], logger)
            primary, secondary = writer.write_workorder_check_result(skipped, [])
            self.finished_signal.emit(primary, secondary)
            return

        # M5: 工单系统比对
        checker = WorkOrderChecker(self.config["workorder"], logger)

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            login_ok = loop.run_until_complete(checker.login())
            if not login_ok:
                logger.error("工单系统登录失败，跳过RPA比对")
                for r in to_query:
                    r["是否已录单"] = "未比对"
                    r["匹配状态"] = "工单系统未连接"
                    r["查询时间戳"] = datetime.now().isoformat(timespec="seconds")
            else:
                try:
                    loop.run_until_complete(checker.search_batch(to_query))
                except Exception as e:
                    logger.error(f"RPA查询异常: {e}", "error")
                finally:
                    loop.run_until_complete(checker.close())
        finally:
            loop.close()

        # 输出 stage2 产物
        from modules.excel_writer import ExcelWriter
        writer = ExcelWriter(self.config["output"]["dir"], logger)
        primary, secondary = writer.write_workorder_check_result(all_rows, to_query, skipped)
        logger.info("=" * 50)
        logger.info(f"[阶段二完成] 工单核对结果: {primary}")

        self.finished_signal.emit(primary, secondary)

    def _run_all(self):
        """一站式: stage1 + stage2 (兼容原行为, 阶段一阶段二连续跑)"""
        import logging
        logger = logging.getLogger("mail_audit")
        logger.setLevel(logging.DEBUG)
        handler = GuiLogHandler(self._log)
        handler.setLevel(logging.INFO)
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
        handler.setFormatter(fmt)
        logger.addHandler(handler)

        logger.info("=" * 50)
        logger.info("[一站式] 邮件 + 工单全流程")
        logger.info("=" * 50)

        all_rows, filtered_mails = self._run_pipeline_stage1(logger)
        if all_rows is None:
            return

        if self._stop:
            return

        # 阶段二手动并行: 直接用 stage1 结果
        from openpyxl import load_workbook
        from modules.workorder_checker import WorkOrderChecker
        from modules.excel_writer import ExcelWriter

        to_query, skipped = WorkOrderChecker.preprocess_rows(all_rows)
        logger.info(f"前置校验: 待RPA {len(to_query)} 条, 跳过 {len(skipped)} 条")

        checker = WorkOrderChecker(self.config["workorder"], logger)
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            login_ok = loop.run_until_complete(checker.login())
            if not login_ok:
                logger.warning("工单系统登录失败")
                for r in all_rows:
                    r["是否已录单"] = "未比对"
                    r["匹配状态"] = "工单系统未连接"
                    r["查询时间戳"] = datetime.now().isoformat(timespec="seconds")
            else:
                try:
                    loop.run_until_complete(checker.search_batch(to_query))
                    # 把查询结果合并回 all_rows
                    queried = {id(r): r for r in to_query}
                    for r in all_rows:
                        q = queried.get(id(r))
                        if q:
                            r["是否已录单"] = q.get("是否已录单", "")
                            r["工单日期"] = q.get("工单日期")
                            r["匹配状态"] = q.get("匹配状态", "")
                            r["工单记录"] = q.get("工单记录", [])
                            r["查询时间戳"] = q.get("查询时间戳")
                except Exception as e:
                    logger.error(f"RPA查询异常: {e}", "error")
                finally:
                    loop.run_until_complete(checker.close())
        finally:
            loop.close()

        writer = ExcelWriter(self.config["output"]["dir"], logger)
        primary, secondary = writer.write_all_outputs(all_rows, filtered_mails)
        logger.info("=" * 50)
        logger.info(f"[一站式完成] 漏单清单: {primary}")
        logger.info(f"[一站式完成] 过滤清单: {secondary}")
        self.finished_signal.emit(primary, secondary)

    def _run_pipeline_stage1(self, logger):
        """阶段一内部 pipeline (M1-M4), 返回 (rows, filtered_mails) 或 None=用户停止"""
        from openpyxl import load_workbook
        ref_cfg = self.config["reference_tables"]

        # 加载附件二: 内部邮箱
        internal_email_path = self.internal_email_path or ref_cfg["internal_emails"]
        internal_emails = self._load_internal_emails(internal_email_path)
        logger.info(f"内部邮箱表已加载: {len(internal_emails)} 条")

        # 加载附件三: 代理邮箱
        agent_email_path = self.agent_email_path or ref_cfg["agent_emails"]
        agent_map = self._load_agent_emails(agent_email_path)
        logger.info(f"代理邮箱表已加载: {len(agent_map)} 条")

        # 加载附件四: 项目名称
        project_table_path = self.project_table_path or ref_cfg["project_names"]
        project_table = self._load_project_names(project_table_path)
        logger.info(f"项目名称表已加载: {len(project_table)} 条")

        # 创建 LLM 客户端
        llm_client = None
        llm_cfg = self.config.get("llm", {})
        if llm_cfg.get("api_key"):
            llm_client = LLMIntentClient(llm_cfg, logger)
            if llm_client.enabled:
                logger.info(f"LLM 已启用: model={llm_cfg.get('model','deepseek-chat')}")
            else:
                logger.warning("LLM 配置异常, 将仅使用规则引擎")
        else:
            logger.info("未配置 LLM, 仅使用规则引擎")

        if self._stop:
            return None

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
            return None

        # M2: 邮件过滤
        from modules.mail_filter import MailFilter
        filt = MailFilter(internal_emails, logger)
        valid_mails, filtered_mails = filt.filter_mails(mails)
        logger.info(f"规则过滤完成: 有效 {len(valid_mails)} 封, 过滤 {len(filtered_mails)} 封")

        if self._stop:
            return None

        # M2.5: LLM 二次意图识别 — 对过滤掉的邮件做 LLM 分类
        if llm_client and llm_client.enabled and filtered_mails:
            logger.info(f"LLM 二次识别开始: 对 {len(filtered_mails)} 封过滤邮件做意图分类...")
            recovered, still_filtered = filt.llm_second_pass(
                filtered_mails, llm_client,
                progress_callback=lambda c, t: self.progress_signal.emit(c, t)
            )
            if recovered:
                logger.info(f"LLM 恢复 {len(recovered)} 封邮件到有效列表")
                valid_mails.extend(recovered)
            filtered_mails = still_filtered
            logger.info(f"LLM 二次识别完成: 有效 {len(valid_mails)} 封, 过滤 {len(filtered_mails)} 封")
        else:
            logger.info("跳过 LLM 二次识别 (未配置或无过滤邮件)")

        if self._stop:
            return None

        # M3: 字段提取 (规则 + LLM 补充 + OCR 图片兜底)
        from modules.field_extractor import FieldExtractor
        ocr_fallback = self.config.get("ocr", {}).get("fallback", True)
        extractor = FieldExtractor(agent_map, project_table, logger, llm_client,
                                   ocr_fallback=ocr_fallback)
        all_rows = []
        for idx, mail in enumerate(valid_mails, 1):
            if self._stop:
                return None
            rows = extractor.extract_fields(mail)
            all_rows.extend(rows)
            logger.info(f"字段提取: 第{idx}/{len(valid_mails)}封 — "
                       f"客户='{rows[0].get('客户','')}', 代理='{rows[0].get('代理','')}'")
            self.progress_signal.emit(idx, len(valid_mails))

        if self._stop:
            return None

        # M4: 项目标准化拆分 — 所有非空项目统一走 normalizer（组合拆分+标准化）
        from modules.project_normalizer import ProjectNormalizer
        normalizer = ProjectNormalizer(project_table, logger)
        normalized_rows = []
        for row in all_rows:
            proj_raw = row.get("项目", "")
            if proj_raw:
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
                if len(split_results) == 1 and split_results[0]["standard_name"]:
                    row["项目"] = split_results[0]["standard_name"]
            normalized_rows.append(row)

        logger.info(f"项目标准化完成: {len(normalized_rows)} 行")

        return normalized_rows, filtered_mails

    def _on_finished_stage2(self, primary_file, secondary_file):
        """阶段二完成信号占位（实际 finished_signal 在 run() 内 emit）"""
        pass

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
        self.load_session()

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

        self.btn_import_internal = QPushButton("内部邮箱表导入")
        self.btn_import_internal.clicked.connect(self.import_internal_table)
        ctrl_layout.addWidget(self.btn_import_internal)
        self.internal_table_label = QLabel("未导入")
        self.internal_table_label.setStyleSheet("color: #999;")
        ctrl_layout.addWidget(self.internal_table_label)

        self.btn_import_project = QPushButton("项目表导入")
        self.btn_import_project.clicked.connect(self.import_project_table)
        ctrl_layout.addWidget(self.btn_import_project)
        self.project_table_label = QLabel("未导入")
        self.project_table_label.setStyleSheet("color: #999;")
        ctrl_layout.addWidget(self.project_table_label)

        ctrl_layout.addStretch()

        # === 运行模式选择 ===
        mode_group = QGroupBox("运行模式")
        mode_layout = QHBoxLayout(mode_group)
        self.radio_all = QRadioButton("一站式 (邮件+工单)")
        self.radio_stage1 = QRadioButton("仅阶段一 (邮件解析)")
        self.radio_stage2 = QRadioButton("仅阶段二 (工单核对)")
        self.radio_all.setChecked(True)
        mode_btn_group = QButtonGroup(self)
        mode_btn_group.addButton(self.radio_all, 0)
        mode_btn_group.addButton(self.radio_stage1, 1)
        mode_btn_group.addButton(self.radio_stage2, 2)
        mode_layout.addWidget(self.radio_all)
        mode_layout.addWidget(self.radio_stage1)
        mode_layout.addWidget(self.radio_stage2)
        ctrl_layout.addWidget(mode_group)

        ctrl_layout.addSpacing(20)

        # === 阶段二输入文件选择 (仅 stage2 模式可见) ===
        self.stage2_input_label = QLabel("阶段二输入: 未选择")
        self.stage2_input_label.setStyleSheet("color: #999;")
        ctrl_layout.addWidget(self.stage2_input_label)
        self.btn_choose_stage2_input = QPushButton("选择阶段一输出文件")
        self.btn_choose_stage2_input.clicked.connect(self.choose_stage2_input)
        ctrl_layout.addWidget(self.btn_choose_stage2_input)
        self.btn_choose_stage2_input.setVisible(False)
        self.stage2_input_label.setVisible(False)
        # 模式切换时显示/隐藏
        self.radio_all.toggled.connect(self._on_mode_changed)
        self.radio_stage1.toggled.connect(self._on_mode_changed)
        self.radio_stage2.toggled.connect(self._on_mode_changed)

        ctrl_layout.addSpacing(10)

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

        # === 凭据配置区 ===
        cred_group = QGroupBox("账号配置 (人员变动时可直接修改并保存)")
        cred_layout = QFormLayout(cred_group)

        # 阿里邮箱
        self.input_email_addr = QLineEdit()
        self.input_email_pwd = QLineEdit()
        self.input_email_pwd.setEchoMode(QLineEdit.Password)
        cred_layout.addRow("阿里邮箱地址:", self.input_email_addr)
        cred_layout.addRow("阿里邮箱密码:", self.input_email_pwd)

        # 工单系统
        self.input_wo_user = QLineEdit()
        self.input_wo_pwd = QLineEdit()
        self.input_wo_pwd.setEchoMode(QLineEdit.Password)
        cred_layout.addRow("工单系统账号:", self.input_wo_user)
        cred_layout.addRow("工单系统密码:", self.input_wo_pwd)

        # 保存按钮
        btn_save_cred = QPushButton("保存配置到 config.yaml")
        btn_save_cred.setStyleSheet(
            "QPushButton { background-color: #2196F3; color: white; padding: 5px 15px; }"
            "QPushButton:hover { background-color: #1976D2; }"
        )
        btn_save_cred.clicked.connect(self.save_credentials)
        cred_layout.addRow(btn_save_cred)

        layout.addWidget(cred_group)

        # === 进度条 ===
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFormat("准备中...")
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

        self.btn_open_stage2_result = QPushButton("查看工单核对结果")
        self.btn_open_stage2_result.setEnabled(False)
        self.btn_open_stage2_result.clicked.connect(lambda: self.open_excel(self.stage2_result_file))
        bottom_layout.addWidget(self.btn_open_stage2_result)

        bottom_layout.addStretch()
        btn_exit = QPushButton("退出")
        btn_exit.clicked.connect(self.close)
        bottom_layout.addWidget(btn_exit)
        layout.addWidget(bottom_widget)

        self.missing_file = ""
        self.filtered_file = ""
        self.stage2_result_file = ""

    def load_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                self.config = yaml.safe_load(f)
        else:
            QMessageBox.critical(self, "错误", "找不到 config.yaml 配置文件!")
            self.config = {}
        self.load_credentials()

    def load_session(self):
        """加载上次保存的时间范围和邮箱汇总表路径"""
        if not os.path.exists(SESSION_FILE):
            return
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
            # 恢复时间范围
            date_from_str = state.get("date_from")
            date_to_str = state.get("date_to")
            if date_from_str:
                parts = date_from_str.split("-")
                self.date_from.setDate(QDate(int(parts[0]), int(parts[1]), int(parts[2])))
            if date_to_str:
                parts = date_to_str.split("-")
                self.date_to.setDate(QDate(int(parts[0]), int(parts[1]), int(parts[2])))
            # 恢复邮箱汇总表路径
            agent_path = state.get("agent_table_path")
            if agent_path and os.path.exists(agent_path):
                self.agent_email_path = agent_path
                fname = os.path.basename(agent_path)
                self.agent_table_label.setText(f"已导入: {fname}")
                self.agent_table_label.setStyleSheet("color: #4CAF50;")
                self.log(f"已恢复上次邮箱汇总表: {fname}")
            # 恢复内部邮箱表路径
            internal_path = state.get("internal_table_path")
            if internal_path and os.path.exists(internal_path):
                self.internal_email_path = internal_path
                fname = os.path.basename(internal_path)
                self.internal_table_label.setText(f"已导入: {fname}")
                self.internal_table_label.setStyleSheet("color: #4CAF50;")
                self.log(f"已恢复上次内部邮箱表: {fname}")
            # 恢复项目表路径
            project_path = state.get("project_table_path")
            if project_path and os.path.exists(project_path):
                self.project_table_path = project_path
                fname = os.path.basename(project_path)
                self.project_table_label.setText(f"已导入: {fname}")
                self.project_table_label.setStyleSheet("color: #4CAF50;")
                self.log(f"已恢复上次项目表: {fname}")
            # 恢复阶段二输入文件路径
            stage2_in = state.get("stage2_input_path")
            if stage2_in and os.path.exists(stage2_in):
                self.stage2_input_path = stage2_in
                fname = os.path.basename(stage2_in)
                self.stage2_input_label.setText(f"阶段二输入: {fname}")
                self.stage2_input_label.setStyleSheet("color: #4CAF50;")
                self.log(f"已恢复上次阶段二输入: {fname}")
            # 恢复运行模式
            mode = state.get("mode", "all")
            if mode == "stage1":
                self.radio_stage1.setChecked(True)
            elif mode == "stage2":
                self.radio_stage2.setChecked(True)
            else:
                self.radio_all.setChecked(True)
            self._on_mode_changed()
        except Exception as e:
            self.log(f"[WARNING] 恢复会话状态失败: {e}")

    def save_session(self):
        """保存当前时间范围和邮箱汇总表路径"""
        if self.radio_stage2.isChecked():
            mode = "stage2"
        elif self.radio_stage1.isChecked():
            mode = "stage1"
        else:
            mode = "all"
        state = {
            "date_from": self.date_from.date().toString("yyyy-MM-dd"),
            "date_to": self.date_to.date().toString("yyyy-MM-dd"),
            "agent_table_path": getattr(self, "agent_email_path", None),
            "internal_table_path": getattr(self, "internal_email_path", None),
            "project_table_path": getattr(self, "project_table_path", None),
            "stage2_input_path": getattr(self, "stage2_input_path", None),
            "mode": mode,
        }
        try:
            with open(SESSION_FILE, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def closeEvent(self, event):
        """窗口关闭时保存会话"""
        self.save_session()
        event.accept()

    def load_credentials(self):
        """从 config 加载凭据到界面输入框"""
        if not self.config:
            return
        email_cfg = self.config.get("email", {})
        self.input_email_addr.setText(email_cfg.get("address", ""))
        self.input_email_pwd.setText(email_cfg.get("password", ""))

        wo_cfg = self.config.get("workorder", {})
        self.input_wo_user.setText(wo_cfg.get("username", ""))
        self.input_wo_pwd.setText(wo_cfg.get("password", ""))

    def save_credentials(self):
        """将界面输入的凭据保存到 config.yaml 并更新内存配置"""
        if not self.config:
            QMessageBox.critical(self, "错误", "配置文件未加载!")
            return

        self.config.setdefault("email", {})
        self.config["email"]["address"] = self.input_email_addr.text().strip()
        self.config["email"]["password"] = self.input_email_pwd.text().strip()

        self.config.setdefault("workorder", {})
        self.config["workorder"]["username"] = self.input_wo_user.text().strip()
        self.config["workorder"]["password"] = self.input_wo_pwd.text().strip()

        config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                yaml.dump(self.config, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
            self.log("凭据已保存到 config.yaml")
            QMessageBox.information(self, "保存成功", "账号配置已保存到 config.yaml")
        except Exception as e:
            self.log(f"[ERROR] 保存配置失败: {e}")
            QMessageBox.critical(self, "保存失败", str(e))

    def choose_stage2_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择阶段一产出的工单待查文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.stage2_input_path = path
            fname = os.path.basename(path)
            self.stage2_input_label.setText(f"阶段二输入: {fname}")
            self.stage2_input_label.setStyleSheet("color: #4CAF50;")
            self.log(f"阶段二输入文件已选: {fname}")
            self.save_session()

    def _on_mode_changed(self):
        is_stage2 = self.radio_stage2.isChecked()
        self.btn_choose_stage2_input.setVisible(is_stage2)
        self.stage2_input_label.setVisible(is_stage2)
        self.btn_import_agent.setVisible(not is_stage2)
        self.agent_table_label.setVisible(not is_stage2)
        self.btn_import_internal.setVisible(not is_stage2)
        self.internal_table_label.setVisible(not is_stage2)
        self.btn_import_project.setVisible(not is_stage2)
        self.project_table_label.setVisible(not is_stage2)
        self.date_from.setVisible(not is_stage2)
        # 找"时间范围"标签隐藏：layout 较复杂，简化只隐藏 date_from/to
        self.date_to.setVisible(not is_stage2)
        self.btn_run.setText("开始运行")
        if is_stage2:
            self.btn_run.setText("开始工单核对")

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
            self.save_session()

    def import_internal_table(self):
        """导入附件二: 内部邮箱表"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择内部邮箱表 Excel 文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.internal_email_path = path
            fname = os.path.basename(path)
            self.internal_table_label.setText(f"已导入: {fname}")
            self.internal_table_label.setStyleSheet("color: #4CAF50;")
            self.log(f"内部邮箱表已导入: {fname}")
            self.save_session()

    def import_project_table(self):
        """导入附件四: 项目名称表"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择项目名称表 Excel 文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.project_table_path = path
            fname = os.path.basename(path)
            self.project_table_label.setText(f"已导入: {fname}")
            self.project_table_label.setStyleSheet("color: #4CAF50;")
            self.log(f"项目名称表已导入: {fname}")
            self.save_session()

    def start_run(self):
        if not self.config:
            QMessageBox.critical(self, "错误", "配置文件未加载!")
            return

        # 模式选择
        if self.radio_stage2.isChecked():
            mode = "stage2"
        elif self.radio_stage1.isChecked():
            mode = "stage1"
        else:
            mode = "all"

        # 阶段二模式必须选输入文件
        if mode == "stage2" and not getattr(self, "stage2_input_path", None):
            QMessageBox.critical(self, "错误", "阶段二模式必须先选择阶段一输出的工单待查文件")
            return

        # 使用界面当前输入的凭据覆盖内存配置
        self.config.setdefault("email", {})
        self.config["email"]["address"] = self.input_email_addr.text().strip()
        self.config["email"]["password"] = self.input_email_pwd.text().strip()
        self.config.setdefault("workorder", {})
        self.config["workorder"]["username"] = self.input_wo_user.text().strip()
        self.config["workorder"]["password"] = self.input_wo_pwd.text().strip()

        # 仅 stage1/all 需要时间范围
        if mode != "stage2":
            date_from = datetime(
                self.date_from.date().year(), self.date_from.date().month(), self.date_from.date().day()
            )
            date_to = datetime(
                self.date_to.date().year(), self.date_to.date().month(), self.date_to.date().day()
            ) + timedelta(days=1)
        else:
            date_from = date_to = None

        agent_path = getattr(self, "agent_email_path", None)
        internal_path = getattr(self, "internal_email_path", None)
        project_path = getattr(self, "project_table_path", None)
        stage2_in = getattr(self, "stage2_input_path", None)

        self.worker = WorkerThread(
            self.config, mode=mode,
            date_from=date_from, date_to=date_to,
            agent_email_path=agent_path,
            internal_email_path=internal_path,
            project_table_path=project_path,
            stage2_input_path=stage2_in,
        )
        self.save_session()
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.on_progress)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.start()

        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_open_missing.setEnabled(False)
        self.btn_open_filtered.setEnabled(False)
        self.btn_open_stage2_result.setEnabled(False)
        self.progress.setValue(0)

    def stop_run(self):
        if self.worker:
            self.worker.stop()
            self.log("用户请求停止运行...")
            self.btn_stop.setEnabled(False)

    def on_progress(self, current, total):
        if total <= 0:
            return
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.progress.setFormat(f"{current} / {total}")

    def on_finished(self, primary_file, secondary_file):
        # 根据当前模式分别放置文件名 + 启用按钮
        if self.radio_stage2.isChecked():
            self.stage2_result_file = primary_file
            self.btn_open_stage2_result.setEnabled(True)
            self.log(f"工单核对结果: {primary_file}")
        elif self.radio_stage1.isChecked():
            self.missing_file = primary_file
            self.filtered_file = secondary_file
            self.btn_open_missing.setEnabled(True)
            self.btn_open_filtered.setEnabled(True)
            self.log(f"阶段一待查清单: {primary_file}")
            self.log(f"阶段一漏单复查: {secondary_file}")
        else:
            # 一站式: primary=漏单清单, secondary=过滤清单
            self.missing_file = primary_file
            self.filtered_file = secondary_file
            self.btn_open_missing.setEnabled(True)
            self.btn_open_filtered.setEnabled(True)
            self.log(f"漏单清单: {primary_file}")
            self.log(f"过滤清单: {secondary_file}")
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.progress.setFormat("完成")
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
