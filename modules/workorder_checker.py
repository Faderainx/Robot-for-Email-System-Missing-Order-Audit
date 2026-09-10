"""M5 工单系统比对模块 — Playwright 半自动登录（验证码人工输入），查询并比对

操作流程:
  1. 登录（自动填账号密码，人工输验证码）
  2. 进入注册工单页面
  3. 逐条数据查询:
     a. 填入公司名称（文本输入）
     b. 选择所属代理（下拉框）
     c. 选择国家（下拉框）
     d. 选择服务项目（下拉框）
     e. 选择状态（下拉框）
     f. 点击查询
     g. 无结果 → 模糊匹配公司名称重新查询
     h. 读取表格结果
  4. 四维模糊比对
"""
import asyncio
import json
import os
import time
from datetime import datetime, timedelta
from typing import List, Dict, Optional

from utils.fuzzy_match import fuzzy_match_pair, normalize_text


# 默认选择器（config.yaml 的 workorder.selectors 段可按键覆盖，无需改代码）
DEFAULT_SELECTORS = {
    "login_account": [
        'input[placeholder="账号"]',
        'input[name="username"]',
        'input[name="account"]',
        'input[type="text"]',
        'input:first-of-type',
    ],
    "login_password": [
        'input[type="password"]',
        'input[placeholder="密码"]',
        'input[name="password"]',
    ],
    "nav": [
        'a:has-text("注册工单")',
        'a:has-text("注册")',
        'li:has-text("注册工单")',
        'li:has-text("注册")',
        'div:has-text("注册工单")',
        'a:has-text("工单")',
        'li:has-text("工单")',
        'a:has-text("订单")',
        '[href*="register"]',
        '[href*="order"]',
    ],
    "dropdown_option": [
        ".el-select-dropdown__item",
        ".ant-select-item",
        ".ant-select-option",
        "li[role='option']",
        "div[role='option']",
        ".dropdown-item",
        ".select-option",
    ],
    "company_input": [
        'input[placeholder*="公司"]',
        'input[placeholder*="客户"]',
        'input[placeholder*="名称"]',
        'input[placeholder*="请输入"]',
        'input[name*="company"]',
        'input[name*="customer"]',
        'input[name*="name"]',
    ],
    "query_button": [
        'button:has-text("查询")',
        'button:has-text("搜索")',
        'button:has-text("查找")',
        'button:has-text("检索")',
        'button[type="submit"]',
        'button.el-button--primary:has-text("查")',
        'button.ant-btn-primary:has-text("查")',
    ],
    "table": [
        ".el-table__body-wrapper table",
        ".el-table table",
        ".ant-table-content table",
        "table.table",
        "table.data-table",
        "table",
    ],
}


class WorkOrderChecker:
    def __init__(self, config: dict, logger=None):
        self.url = config["url"]
        self.username = config["username"]
        self.password = config["password"]
        self.timeout = config.get("timeout", 30)
        self.date_tolerance = config.get("date_tolerance_days", 60)
        self.customer_threshold = config.get("customer_threshold", 85)
        self.selectors = {**DEFAULT_SELECTORS, **(config.get("selectors") or {})}
        # 断点续查 / 风控
        self.query_interval = float(config.get("query_interval_seconds", 3))
        self.cache_path = config.get("query_cache_path", "storage/query_cache.json")
        self.cache_ttl_days = int(config.get("query_cache_ttl_days", 7))
        self.max_total_seconds = float(config.get("max_total_seconds", 0))  # 0=不限
        self._session_start_ts = None
        self.logger = logger
        self._playwright = None
        self._browser = None
        self._page = None
        self._logged_in = False
        # 加载缓存
        self._query_cache: Dict[str, dict] = self._load_cache()

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    async def _ensure_browser(self):
        if self._page is not None:
            return
        from playwright.async_api import async_playwright
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(headless=False)
        self._page = await self._browser.new_page()
        self._page.set_default_timeout(self.timeout * 1000)

    async def login(self) -> bool:
        if self._logged_in:
            return True
        await self._ensure_browser()
        try:
            self._log(f"导航到工单系统: {self.url}")
            await self._page.goto(self.url, wait_until="networkidle")
            await self._page.wait_for_load_state("networkidle")

            login_url = self._page.url
            self._log(f"当前登录页: {login_url}")
            await self._page.screenshot(path="output/workorder_login.png", full_page=False)

            # 填账号
            account_filled = False
            for sel in self.selectors["login_account"]:
                try:
                    el = await self._page.wait_for_selector(sel, timeout=3000)
                    if el:
                        await el.fill("")
                        await el.type(self.username)
                        account_filled = True
                        self._log(f"账号已填写 (selector={sel})")
                        break
                except Exception:
                    continue

            if not account_filled:
                self._log("未找到账号输入框", "error")
                return False

            # 填密码
            password_filled = False
            for sel in self.selectors["login_password"]:
                try:
                    el = await self._page.wait_for_selector(sel, timeout=3000)
                    if el:
                        await el.fill("")
                        await el.type(self.password)
                        password_filled = True
                        self._log(f"密码已填写 (selector={sel})")
                        break
                except Exception:
                    continue

            if not password_filled:
                self._log("未找到密码输入框", "error")
                return False

            self._log("=" * 50)
            self._log("账号密码已自动填写")
            self._log(">>> 请在浏览器中输入验证码并点击登录按钮 <<<")
            self._log("=" * 50)
            self._log("等待登录完成... (最长等待 120 秒)")

            try:
                await self._page.wait_for_url(
                    lambda url: "/login" not in url,
                    timeout=120000,
                )
                self._log("检测到页面跳转，登录成功")
                await self._page.wait_for_load_state("networkidle")
                await asyncio.sleep(2)
                self._logged_in = True
                await self._page.screenshot(path="output/workorder_after_login.png", full_page=False)
                page_text = await self._page.evaluate("() => document.body.innerText.substring(0, 500)")
                self._log(f"页面文本预览: {page_text[:200]}")
                return True
            except Exception:
                current = self._page.url
                if "/login" not in current and "workbench" in current:
                    self._log("已在工单系统页面，登录成功")
                    self._logged_in = True
                    return True
                self._log(f"登录超时，当前URL: {current}", "error")
                return False
        except Exception as e:
            self._log(f"工单系统登录失败: {e}", "error")
            return False

    async def _navigate_to_order_list(self):
        """导航到注册工单页面"""
        # 尝试找到"注册工单"入口
        nav_selectors = self.selectors["nav"]

        for sel in nav_selectors:
            try:
                el = await self._page.query_selector(sel)
                if el:
                    text = await el.inner_text()
                    if any(kw in text for kw in ["注册工单", "注册", "工单", "订单"]):
                        await el.click()
                        await self._page.wait_for_load_state("networkidle")
                        await asyncio.sleep(2)
                        self._log(f"导航到注册工单页面 (点击: {text[:30]})")
                        await self._page.screenshot(path="output/workorder_reg_page.png", full_page=False)
                        return
            except Exception:
                continue

        self._log("未找到注册工单导航入口，尝试在当前页操作", "warning")
        await self._page.screenshot(path="output/workorder_nav_fail.png", full_page=False)

    async def _select_dropdown(self, field_name: str, value: str) -> bool:
        """
        通用下拉框选择: 找到标签旁边的下拉框，点击展开，选对应选项
        field_name: 字段中文名（如"所属代理""国家""服务项目""状态"）
        value: 要选的值
        """
        if not value:
            self._log(f"{field_name} 值为空，跳过选择")
            return True

        value_norm = normalize_text(value)

        # 策略1: 找到字段标签，然后找旁边的 select 或下拉组件
        label_selectors = [
            f'label:has-text("{field_name}")',
            f'span:has-text("{field_name}")',
            f'div:has-text("{field_name}") >> nth=0',
        ]

        for label_sel in label_selectors:
            try:
                label = await self._page.query_selector(label_sel)
                if not label:
                    continue

                # 找标签附近的 select 元素
                select = await label.evaluate_handle(
                    "el => el.parentElement?.querySelector('select, .el-select, .ant-select, [class*=\"select\"], [class*=\"dropdown\"]')"
                )
                if select:
                    select_el = select.as_element()
                    if select_el:
                        # 如果是原生 select
                        tag = await select_el.evaluate("el => el.tagName.toLowerCase()")
                        if tag == "select":
                            await select_el.select_option(label=value)
                            self._log(f"{field_name} 选择: {value} (原生select)")
                            return True
                        # 如果是自定义下拉组件，点击展开
                        await select_el.click()
                        await asyncio.sleep(0.5)
                        # 找选项
                        option = await self._find_dropdown_option(value_norm)
                        if option:
                            await option.click()
                            await asyncio.sleep(0.3)
                            self._log(f"{field_name} 选择: {value} (自定义下拉)")
                            return True
            except Exception:
                continue

        # 策略2: 直接找页面上所有下拉框，按顺序匹配
        all_selects = await self._page.query_selector_all(
            '.el-select, .ant-select, select, [class*="select"]:not([class*="selected"]):not(input)'
        )
        for sel_el in all_selects:
            try:
                tag = await sel_el.evaluate("el => el.tagName.toLowerCase()")
                if tag == "select":
                    options = await sel_el.query_selector_all("option")
                    for opt in options:
                        opt_text = await opt.inner_text()
                        if value_norm in normalize_text(opt_text) or normalize_text(opt_text) in value_norm:
                            await sel_el.select_option(label=opt_text.strip())
                            self._log(f"{field_name} 选择: {opt_text.strip()} (原生select遍历)")
                            return True
                else:
                    # 自定义下拉，点击展开
                    await sel_el.click()
                    await asyncio.sleep(0.5)
                    option = await self._find_dropdown_option(value_norm)
                    if option:
                        await option.click()
                        await asyncio.sleep(0.3)
                        self._log(f"{field_name} 选择: {value} (自定义下拉遍历)")
                        return True
                    # 关闭下拉
                    await self._page.keyboard.press("Escape")
                    await asyncio.sleep(0.2)
            except Exception:
                continue

        self._log(f"{field_name} 未找到匹配选项: {value}", "warning")
        return False

    async def _find_dropdown_option(self, value_norm: str):
        """在下拉框展开后，找到匹配的选项"""
        option_selectors = self.selectors["dropdown_option"]
        for sel in option_selectors:
            options = await self._page.query_selector_all(sel)
            for opt in options:
                try:
                    opt_text = await opt.inner_text()
                    opt_norm = normalize_text(opt_text.strip())
                    if value_norm in opt_norm or opt_norm in value_norm:
                        return opt
                except Exception:
                    continue
        return None

    async def _fill_company_name(self, company: str) -> bool:
        """填入公司名称（文本输入框）"""
        # 找到公司名称输入框
        company_selectors = self.selectors["company_input"]

        for sel in company_selectors:
            try:
                el = await self._page.wait_for_selector(sel, timeout=3000)
                if el:
                    await el.fill("")
                    await el.type(company)
                    self._log(f"公司名称已填入: {company} (selector={sel})")
                    return True
            except Exception:
                continue

        # 兜底: 找所有 text input
        all_inputs = await self._page.query_selector_all('input[type="text"], input:not([type])')
        for inp in all_inputs:
            try:
                placeholder = await inp.get_attribute("placeholder") or ""
                if any(kw in placeholder for kw in ["公司", "客户", "名称", "请输入"]):
                    await inp.fill("")
                    await inp.type(company)
                    self._log(f"公司名称已填入: {company} (兜底selector)")
                    return True
            except Exception:
                continue

        self._log(f"未找到公司名称输入框", "error")
        return False

    async def _clear_inputs(self):
        """清空所有文本输入框，为下一条数据做准备"""
        all_inputs = await self._page.query_selector_all('input[type="text"], input:not([type]), textarea')
        for inp in all_inputs:
            try:
                await inp.fill("")
            except Exception:
                continue

        # 重置下拉框
        all_selects = await self._page.query_selector_all('.el-select, .ant-select')
        for sel in all_selects:
            try:
                # 找清空按钮
                clear_btn = await sel.query_selector('.el-select__clear, .ant-select-clear, [class*="clear"]')
                if clear_btn:
                    await clear_btn.click()
                    await asyncio.sleep(0.2)
            except Exception:
                continue

    async def _click_query_button(self) -> bool:
        """点击查询按钮"""
        btn_selectors = self.selectors["query_button"]
        for sel in btn_selectors:
            try:
                btn = await self._page.wait_for_selector(sel, timeout=3000)
                if btn:
                    await btn.click()
                    self._log(f"点击查询按钮: {sel}")
                    await self._page.wait_for_load_state("networkidle")
                    await asyncio.sleep(2)
                    return True
            except Exception:
                continue
        self._log("未找到查询按钮", "error")
        return False

    async def search_one(
        self,
        company: str,
        agent: str = "",
        country: str = "",
        service_item: str = "",
        status: str = "",
    ) -> List[Dict]:
        """
        查询单条数据:
        1. 填公司名称
        2. 选所属代理（下拉框）
        3. 选国家（下拉框）
        4. 选服务项目（下拉框）
        5. 选状态（下拉框）
        6. 点查询
        7. 无结果 → 模糊匹配公司名重查
        """
        if not self._logged_in:
            success = await self.login()
            if not success:
                return []

        # 断点续查：命中缓存直接返回，不再触发 RPA
        project_norm = self._compose_project_for_query(agent, country, service_item)
        cached = self.get_cached_result(agent, company, project_norm)
        if cached is not None:
            self._log(f"工单缓存命中: 公司={company}, 代理={agent}, 项目={project_norm} ({len(cached)} 条)")
            return self._orders_from_jsonable(cached)

        workorders = []

        try:
            # 清空上次查询条件
            await self._clear_inputs()
            await asyncio.sleep(0.3)

            # 1. 填公司名称
            filled = await self._fill_company_name(company)
            if not filled:
                return []

            # 2. 选所属代理
            if agent:
                await self._select_dropdown("所属代理", agent)
                await asyncio.sleep(0.3)

            # 3. 选国家
            if country:
                await self._select_dropdown("国家", country)
                await asyncio.sleep(0.3)

            # 4. 选服务项目
            if service_item:
                await self._select_dropdown("服务项目", service_item)
                await asyncio.sleep(0.3)

            # 5. 选状态
            if status:
                await self._select_dropdown("状态", status)
                await asyncio.sleep(0.3)

            # 6. 点查询
            await self._page.screenshot(path="output/workorder_before_query.png", full_page=False)
            clicked = await self._click_query_button()
            if not clicked:
                return []

            await self._page.screenshot(path="output/workorder_after_query.png", full_page=False)

            # 7. 读取结果
            workorders = await self._extract_table_data()
            self._log(f"精确查询结果: {len(workorders)} 条 (公司={company})")

            # 8. 无结果 → 模糊匹配公司名称（只改公司名，其他不动）
            if not workorders and len(company) > 2:
                self._log(f"精确查询无结果，尝试模糊匹配公司名称...")

                # 尝试用公司名的关键词搜索
                # 去掉常见后缀，取核心词
                fuzzy_terms = self._get_fuzzy_company_terms(company)

                for term in fuzzy_terms:
                    if term == company:
                        continue
                    # 只清空公司名输入框，其他下拉不动
                    company_inputs = await self._page.query_selector_all(
                        'input[placeholder*="公司"], input[placeholder*="客户"], input[placeholder*="名称"]'
                    )
                    for ci in company_inputs:
                        try:
                            await ci.fill("")
                            await ci.type(term)
                        except Exception:
                            continue

                    await self._click_query_button()
                    workorders = await self._extract_table_data()
                    if workorders:
                        self._log(f"模糊匹配命中: '{term}' → {len(workorders)} 条")
                        break

                if not workorders:
                    self._log(f"模糊匹配也无结果: {company}", "warning")

        except Exception as e:
            self._log(f"工单查询失败: {e}", "error")

        # 写缓存（含异常路径，避免下次重跑同样失败）
        self.save_to_cache(agent, company, project_norm, self._orders_to_jsonable(workorders))
        return workorders

    def _compose_project_for_query(self, agent: str, country: str, service_item: str) -> str:
        """country+service_item 拼成一个稳定的项目键供缓存去重"""
        return f"{country or ''}_{service_item or ''}"

    def _get_fuzzy_company_terms(self, company: str) -> List[str]:
        """生成公司名称的模糊匹配搜索词（从长到短）"""
        terms = []
        # 去掉常见后缀
        suffixes = ["有限公司", "有限责任公司", "股份公司", "股份有限公司", "集团", "公司"]
        cleaned = company
        for s in suffixes:
            if cleaned.endswith(s):
                cleaned = cleaned[:-len(s)]
                break
        if cleaned and cleaned != company:
            terms.append(cleaned)

        # 取公司名中的关键词（去掉"市""省"等）
        import re
        core = re.sub(r"[省市特区]", "", cleaned)
        core = core.strip()
        if core and core not in terms:
            terms.append(core)

        # 如果还长，取前4个字
        if len(cleaned) > 4:
            terms.append(cleaned[:4])

        # 原始名放最后
        terms.append(company)
        return terms

    async def _extract_table_data(self) -> List[Dict]:
        """从页面表格中提取工单数据"""
        workorders = []

        table_selectors = self.selectors["table"]

        for sel in table_selectors:
            try:
                tables = await self._page.query_selector_all(sel)
                for table in tables:
                    rows = await table.query_selector_all("tr")
                    if len(rows) <= 1:
                        continue

                    header_cells = await rows[0].query_selector_all("th, td")
                    headers = []
                    for cell in header_cells:
                        text = await cell.inner_text()
                        headers.append(text.strip())

                    if not headers:
                        headers = [f"col_{i}" for i in range(20)]

                    self._log(f"表格: {len(headers)} 列, {len(rows)-1} 行, 表头={headers[:8]}")

                    for row in rows[1:]:
                        cells = await row.query_selector_all("td")
                        if not cells:
                            continue
                        row_data = {}
                        for idx, cell in enumerate(cells):
                            text = await cell.inner_text()
                            key = headers[idx] if idx < len(headers) else f"col_{idx}"
                            row_data[key] = text.strip()
                        if any(v for v in row_data.values()):
                            workorders.append(row_data)

                if workorders:
                    break

            except Exception as e:
                self._log(f"选择器 {sel} 表格提取失败: {e}", "warning")
                continue

        if not workorders:
            try:
                all_rows = await self._page.query_selector_all("tr")
                if len(all_rows) > 1:
                    self._log(f"找到 {len(all_rows)} 个 tr (通用模式)")
                    header_cells = await all_rows[0].query_selector_all("th, td")
                    headers = []
                    for cell in header_cells:
                        text = await cell.inner_text()
                        headers.append(text.strip())
                    if not headers:
                        headers = [f"col_{i}" for i in range(20)]

                    for row in all_rows[1:200]:
                        cells = await row.query_selector_all("td")
                        if not cells:
                            continue
                        row_data = {}
                        for idx, cell in enumerate(cells):
                            text = await cell.inner_text()
                            key = headers[idx] if idx < len(headers) else f"col_{idx}"
                            row_data[key] = text.strip()
                        if any(v for v in row_data.values()):
                            workorders.append(row_data)
            except Exception as e:
                self._log(f"通用 tr 提取失败: {e}", "error")

        return workorders

    def _extract_country_from_project(self, project: str) -> str:
        """从项目名称中提取国家"""
        countries = [
            "德国", "法国", "意大利", "西班牙", "荷兰", "波兰",
            "瑞典", "比利时", "爱尔兰", "葡萄牙", "奥地利",
            "英国", "匈牙利", "捷克", "丹麦", "芬兰", "挪威",
            "罗马尼亚", "保加利亚", "希腊", "克罗地亚",
        ]
        for c in countries:
            if c in project:
                return c
        return ""

    def _extract_service_item_from_project(self, project: str) -> str:
        """从项目名称中提取服务项目类型"""
        if "WEEE" in project.upper():
            return "WEEE"
        if "电池" in project:
            return "电池法"
        if "包装" in project:
            return "包装法"
        if "EPR" in project.upper():
            return "EPR"
        if "一次性塑料" in project:
            return "一次性塑料"
        return ""

    async def search_batch(self, mail_rows: List[Dict]) -> List[Dict]:
        """
        批量查询: 逐条数据进入注册工单页面，填条件，查询，读取结果
        每条数据只与本行查询结果比对（按行隔离，防跨行串扰）
        返回: 已写入 是否已录单/匹配状态 字段的行列表
        风控: 条目间间隔 query_interval_seconds；总耗时超 max_total_seconds 主动终止
        断点续查: 已查询 (代理+公司+项目) 三联键命中缓存直接复用，不再走 RPA
        """
        if not self._logged_in:
            success = await self.login()
            if not success:
                return []

        self._session_start_ts = time.monotonic()
        # 导航到注册工单页面
        await self._navigate_to_order_list()

        all_results = []
        total = len(mail_rows)
        terminated = False

        for idx, row in enumerate(mail_rows, 1):
            # 总超时检查
            if self._check_total_timeout():
                self._log(
                    f"⛔ 已达到总超时 {self.max_total_seconds}s, 终止于 {idx-1}/{total}, "
                    f"剩余 {total - idx + 1} 条下次续跑",
                    "error",
                )
                terminated = True
                break

            if not row.get("客户"):
                continue

            company = row.get("客户", "")
            agent = row.get("代理", "")
            project = row.get("项目", "")
            country = self._extract_country_from_project(project)
            service_item = self._extract_service_item_from_project(project)

            self._log(f"工单查询 {idx}/{total}: 公司={company}, 代理={agent}, "
                       f"国家={country}, 项目={service_item}")

            orders = await self.search_one(
                company=company,
                agent=agent,
                country=country,
                service_item=service_item,
            )

            # 按行隔离：本行只与本行查询结果比对，防止跨行串扰掩盖漏单
            self.match_records([row], orders)
            all_results.append(row)
            self._log(f"  → 查到 {len(orders)} 条工单")

            # 条目间间隔（最后一条不 sleep）
            if idx < total and self.query_interval > 0:
                await asyncio.sleep(self.query_interval)

        if terminated:
            elapsed = time.monotonic() - (self._session_start_ts or time.monotonic())
            self._log(f"批次查询被超时终止, 已处理 {len(all_results)}/{total} 条 (耗时 {elapsed:.0f}s)")

        return all_results

    def match_records(
        self,
        mail_rows: List[Dict],
        workorders: List[Dict],
    ) -> List[Dict]:
        """四维模糊比对: 日期 ± 代理 + 客户 + 项目"""
        results = []
        query_timestamp = datetime.now()

        for row in mail_rows:
            mail_date = row.get("date")
            mail_agent = row.get("代理", "")
            mail_customer = row.get("客户", "")
            mail_project = row.get("项目", "")

            matched_orders = []

            for wo in workorders:
                wo_date = self._parse_wo_date(wo)
                date_ok = False
                date_anomaly = False
                if mail_date and wo_date:
                    delta = abs((mail_date - wo_date).days)
                    if delta <= self.date_tolerance:
                        date_ok = True
                    else:
                        date_anomaly = True
                elif not mail_date or not wo_date:
                    date_ok = True

                wo_agent = wo.get("代理", "") or wo.get("agent", "") or wo.get("所属代理", "") or ""
                agent_ok = self._match_agent(mail_agent, wo_agent)

                wo_customer = wo.get("客户", "") or wo.get("customer", "") or wo.get("公司", "") or wo.get("公司名称", "") or ""
                customer_ok = fuzzy_match_pair(mail_customer, wo_customer, self.customer_threshold)

                wo_project = wo.get("项目", "") or wo.get("project", "") or wo.get("服务项目", "") or ""
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
        date_str = (wo.get("日期") or wo.get("工单日期") or wo.get("date")
                    or wo.get("创建时间") or wo.get("注册时间") or "")
        if not date_str:
            return None
        formats = [
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d",
            "%Y年%m月%d日", "%d/%m/%Y",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str[:19] if len(date_str) > 19 else date_str, fmt)
            except ValueError:
                continue
        return None

    def _match_agent(self, mail_agent: str, wo_agent: str) -> bool:
        if not mail_agent or not wo_agent:
            return False
        if fuzzy_match_pair(mail_agent, wo_agent, 80):
            return True
        if mail_agent in wo_agent or wo_agent in mail_agent:
            return True
        return False

    def _match_project(self, mail_project: str, wo_project: str) -> bool:
        if not mail_project or not wo_project:
            return False
        if fuzzy_match_pair(mail_project, wo_project, 85):
            return True
        return False

    async def close(self):
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        self._page = None
        self._browser = None
        self._playwright = None
        self._logged_in = False
        self._session_start_ts = None

    # ---------------- 断点续查 / 请求间隔 ----------------

    @staticmethod
    def _cache_key(agent: str, company: str, project: str) -> str:
        """规范化键：agent|company|project (繁简统一、去空格)"""
        a = normalize_text(agent or "").strip()
        c = normalize_text(company or "").strip()
        p = normalize_text(project or "").strip()
        return f"{a}|{c}|{p}"

    def _load_cache(self) -> Dict[str, dict]:
        if not self.cache_path or not os.path.exists(self.cache_path):
            return {}
        try:
            with open(self.cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                self._log(f"已加载查询缓存: {len(data)} 条 ({self.cache_path})")
                return data
        except Exception as e:
            self._log(f"查询缓存加载失败，忽略: {e}", "warning")
        return {}

    def _save_cache(self):
        if not self.cache_path:
            return
        try:
            os.makedirs(os.path.dirname(self.cache_path) or ".", exist_ok=True)
            tmp = self.cache_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self._query_cache, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.cache_path)
        except Exception as e:
            self._log(f"查询缓存写入失败: {e}", "warning")

    def _is_expired(self, entry: dict) -> bool:
        if self.cache_ttl_days <= 0:
            return False
        ts = entry.get("timestamp")
        if not ts:
            return True
        try:
            age = datetime.now() - datetime.fromisoformat(ts)
            return age.days >= self.cache_ttl_days
        except Exception:
            return True

    def get_cached_result(self, agent: str, company: str, project: str) -> Optional[List[Dict]]:
        """命中且未过期返回缓存 orders，否则 None"""
        key = self._cache_key(agent, company, project)
        entry = self._query_cache.get(key)
        if not entry:
            return None
        if self._is_expired(entry):
            self._query_cache.pop(key, None)
            return None
        return entry.get("orders", [])

    def save_to_cache(self, agent: str, company: str, project: str, orders: List[Dict]):
        key = self._cache_key(agent, company, project)
        self._query_cache[key] = {
            "orders": orders,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }
        self._save_cache()

    def _check_total_timeout(self) -> bool:
        """总超时检查：True=已超时需终止"""
        if self.max_total_seconds <= 0 or self._session_start_ts is None:
            return False
        return (time.monotonic() - self._session_start_ts) >= self.max_total_seconds

    @staticmethod
    def _orders_to_jsonable(orders: List[Dict]) -> List[Dict]:
        """datetime 转 ISO 字符串以便 JSON 化"""
        out = []
        for wo in orders:
            row = {}
            for k, v in wo.items():
                if isinstance(v, datetime):
                    row[k] = v.isoformat()
                else:
                    row[k] = v
            out.append(row)
        return out

    @staticmethod
    def _orders_from_jsonable(orders: List[Dict]) -> List[Dict]:
        """ISO 字符串还原为 datetime"""
        out = []
        for wo in orders:
            row = {}
            for k, v in wo.items():
                if isinstance(v, str) and len(v) >= 10 and v[4] == "-" and v[7] == "-":
                    try:
                        row[k] = datetime.fromisoformat(v)
                        continue
                    except ValueError:
                        pass
                row[k] = v
            out.append(row)
        return out

    # ---------------- 前置校验 ----------------

    @staticmethod
    def preprocess_rows(
        mail_rows: List[Dict],
        skip_threshold: str = "需人工确认",
    ) -> tuple:
        """
        阶段二前置预处理：
          1. 跳过「代理为空 且 置信度=需人工确认」的条目（不执行RPA查询，单独返回告警）
          2. 按 (公司, 项目) 去重 — 同组合只查一次，节省 RPA 时间
        返回: (rows_to_query, skipped_rows)
        """
        rows_to_query: List[Dict] = []
        skipped_rows: List[Dict] = []
        seen = set()

        for r in mail_rows:
            agent = (r.get("代理") or "").strip()
            confidence = (r.get("置信度") or "").strip()

            # 跳过：代理空 + 置信度=需人工确认
            if not agent and confidence == skip_threshold:
                skipped_rows.append({
                    **r,
                    "_skip_reason": "代理空且置信度=需人工确认, 阶段二不查RPA, 请人工补全后手动导入",
                })
                continue

            # 去重：(公司, 项目)
            key = (normalize_text(r.get("客户", "")), normalize_text(r.get("项目", "")))
            if key in seen:
                skipped_rows.append({**r, "_skip_reason": "公司-项目重复, 与前条合并查询"})
                continue
            seen.add(key)
            rows_to_query.append(r)

        return rows_to_query, skipped_rows
