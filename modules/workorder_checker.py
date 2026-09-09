"""M5 工单系统比对模块 — Playwright 自动登录工单系统，查询并比对"""
import asyncio
from datetime import datetime, timedelta
from typing import List, Dict, Optional

from utils.fuzzy_match import fuzzy_match_pair, normalize_text


class WorkOrderChecker:
    def __init__(self, config: dict, logger=None):
        self.url = config["url"]
        self.username = config["username"]
        self.password = config["password"]
        self.timeout = config.get("timeout", 30)
        self.date_tolerance = config.get("date_tolerance_days", 60)
        self.customer_threshold = config.get("customer_threshold", 85)
        self.logger = logger
        self._playwright = None
        self._browser = None
        self._page = None

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    async def _ensure_browser(self):
        """启动 Playwright 浏览器"""
        if self._page is not None:
            return

        from playwright.async_api import async_playwright
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(headless=False)
        self._page = await self._browser.new_page()
        self._page.set_default_timeout(self.timeout * 1000)

    async def login(self) -> bool:
        """登录工单系统"""
        await self._ensure_browser()
        try:
            self._log(f"导航到工单系统: {self.url}")
            await self._page.goto(self.url, wait_until="networkidle")

            # 等待页面加载，尝试定位登录表单
            await self._page.wait_for_load_state("networkidle")

            # 尝试找到用户名输入框
            username_selectors = [
                'input[name="username"]',
                'input[name="userName"]',
                'input[name="account"]',
                'input[type="text"]',
                'input[placeholder*="用户"]',
                'input[placeholder*="账号"]',
                'input[placeholder*="user" i]',
            ]
            username_input = None
            for sel in username_selectors:
                try:
                    el = await self._page.wait_for_selector(sel, timeout=3000)
                    if el:
                        username_input = el
                        break
                except Exception:
                    continue

            if not username_input:
                self._log("未找到用户名输入框，可能已登录或页面结构不同", "warning")
                return True

            await username_input.fill(self.username)

            # 找密码输入框
            password_selectors = [
                'input[name="password"]',
                'input[name="passwd"]',
                'input[type="password"]',
            ]
            password_input = None
            for sel in password_selectors:
                try:
                    el = await self._page.wait_for_selector(sel, timeout=3000)
                    if el:
                        password_input = el
                        break
                except Exception:
                    continue

            if password_input:
                await password_input.fill(self.password)

            # 找登录按钮
            login_selectors = [
                'button[type="submit"]',
                'button:has-text("登录")',
                'button:has-text("登 录")',
                'input[type="submit"]',
                'button:has-text("Login")',
            ]
            for sel in login_selectors:
                try:
                    btn = await self._page.wait_for_selector(sel, timeout=3000)
                    if btn:
                        await btn.click()
                        break
                except Exception:
                    continue

            await self._page.wait_for_load_state("networkidle")
            self._log("工单系统登录成功")
            return True

        except Exception as e:
            self._log(f"工单系统登录失败: {e}", "error")
            return False

    async def search_workorders(self, customer: str, agent: str = "") -> List[Dict]:
        """
        在工单系统中搜索工单记录
        返回: [{"工单日期": datetime, "代理": str, "客户": str, "项目": str, "工单号": str}, ...]
        """
        await self._ensure_browser()
        if self._page is None:
            await self.login()

        workorders = []
        try:
            # 尝试找到搜索/查询输入框
            search_selectors = [
                'input[name="keyword"]',
                'input[name="search"]',
                'input[placeholder*="搜索"]',
                'input[placeholder*="查询"]',
                'input[placeholder*="客户"]',
                'input[type="search"]',
            ]

            search_input = None
            for sel in search_selectors:
                try:
                    el = await self._page.wait_for_selector(sel, timeout=3000)
                    if el:
                        search_input = el
                        break
                except Exception:
                    continue

            if search_input:
                await search_input.fill(customer)
                # 找搜索按钮
                search_btn_selectors = [
                    'button:has-text("搜索")',
                    'button:has-text("查询")',
                    'button:has-text("Search")',
                    'button[type="submit"]',
                ]
                for sel in search_btn_selectors:
                    try:
                        btn = await self._page.wait_for_selector(sel, timeout=2000)
                        if btn:
                            await btn.click()
                            break
                    except Exception:
                        continue

                await self._page.wait_for_load_state("networkidle")
                await asyncio.sleep(1)

            # 尝试从页面表格中提取工单数据
            workorders = await self._extract_table_data()

        except Exception as e:
            self._log(f"工单搜索失败: {e}", "error")

        return workorders

    async def _extract_table_data(self) -> List[Dict]:
        """从页面表格中提取工单数据"""
        workorders = []
        try:
            # 尝试找到表格
            table_selectors = ["table", ".el-table", ".ant-table", ".table"]
            for sel in table_selectors:
                try:
                    table = await self._page.query_selector(sel)
                    if table:
                        rows = await table.query_selector_all("tr")
                        if len(rows) > 1:
                            # 提取表头
                            header_cells = await rows[0].query_selector_all("th, td")
                            headers = []
                            for cell in header_cells:
                                text = await cell.inner_text()
                                headers.append(text.strip())

                            # 提取数据行
                            for row in rows[1:]:
                                cells = await row.query_selector_all("td")
                                if not cells:
                                    continue
                                row_data = {}
                                for idx, cell in enumerate(cells):
                                    text = await cell.inner_text()
                                    key = headers[idx] if idx < len(headers) else f"col_{idx}"
                                    row_data[key] = text.strip()
                                if row_data:
                                    workorders.append(row_data)
                            break
                except Exception:
                    continue
        except Exception as e:
            self._log(f"表格数据提取失败: {e}", "error")

        return workorders

    def match_records(
        self,
        mail_rows: List[Dict],
        workorders: List[Dict],
    ) -> List[Dict]:
        """
        四维模糊比对
        mail_rows: 邮件-项目明细行
        workorders: 工单记录
        返回: 每行附加比对结果
        """
        results = []
        query_timestamp = datetime.now()

        for row in mail_rows:
            mail_date = row.get("date")
            mail_agent = row.get("代理", "")
            mail_customer = row.get("客户", "")
            mail_project = row.get("项目", "")

            matched_orders = []

            for wo in workorders:
                # W4a: 日期比对
                wo_date = self._parse_wo_date(wo)
                date_ok = False
                date_anomaly = False
                if mail_date and wo_date:
                    delta = abs((mail_date - wo_date).days)
                    if delta <= self.date_tolerance:
                        date_ok = True
                    else:
                        date_anomaly = True

                # W4b: 代理比对
                wo_agent = wo.get("代理", "") or wo.get("agent", "") or ""
                agent_ok = self._match_agent(mail_agent, wo_agent)

                # W4c: 客户比对
                wo_customer = wo.get("客户", "") or wo.get("customer", "") or wo.get("公司", "") or ""
                customer_ok = fuzzy_match_pair(mail_customer, wo_customer, self.customer_threshold)

                # W4d: 项目比对
                wo_project = wo.get("项目", "") or wo.get("project", "") or ""
                project_ok = self._match_project(mail_project, wo_project)

                if date_ok and agent_ok and customer_ok and project_ok:
                    matched_orders.append({
                        "工单日期": wo_date,
                        "工单代理": wo_agent,
                        "工单客户": wo_customer,
                        "工单项目": wo_project,
                        "日期异常": date_anomaly,
                    })

            if matched_orders:
                row["是否已录单"] = "是"
                row["工单日期"] = matched_orders[0]["工单日期"]
                row["匹配状态"] = "精确匹配" if len(matched_orders) == 1 else "多匹配-人工排查"
                row["工单记录"] = matched_orders
                row["查询时间戳"] = query_timestamp
            else:
                row["是否已录单"] = "否"
                row["工单日期"] = None
                row["匹配状态"] = "漏单"
                row["工单记录"] = []
                row["查询时间戳"] = query_timestamp

            results.append(row)

        return results

    def _parse_wo_date(self, wo: Dict) -> Optional[datetime]:
        """解析工单日期"""
        date_str = wo.get("日期") or wo.get("工单日期") or wo.get("date") or wo.get("创建时间") or ""
        if not date_str:
            return None
        formats = [
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%Y/%m/%d",
            "%Y年%m月%d日",
            "%d/%m/%Y",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str[:19] if len(date_str) > 19 else date_str, fmt)
            except ValueError:
                continue
        return None

    def _match_agent(self, mail_agent: str, wo_agent: str) -> bool:
        """代理匹配"""
        if not mail_agent or not wo_agent:
            return False
        if fuzzy_match_pair(mail_agent, wo_agent, 80):
            return True
        # 检查简称
        if mail_agent in wo_agent or wo_agent in mail_agent:
            return True
        return False

    def _match_project(self, mail_project: str, wo_project: str) -> bool:
        """项目匹配"""
        if not mail_project or not wo_project:
            return False
        if fuzzy_match_pair(mail_project, wo_project, 85):
            return True
        return False

    async def close(self):
        """关闭浏览器"""
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        self._page = None
        self._browser = None
        self._playwright = None
