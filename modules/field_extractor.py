"""M3 字段提取模块 — 表驱动匹配为主, LLM 补充提取为辅

流程:
  E1 规则提取: 需求(正则) → 代理(邮箱查表) → 客户(主题→正文→附件) → 项目(附件四扫描)
  E2 LLM 补充: 对规则提取后仍有空字段的邮件, 调 LLM 补充提取
     - 仅对缺失字段做补充, 不覆盖规则已提取的结果
     - LLM 提取的项目需与附件四做二次匹配标准化
     - LLM 失败 → 保持规则结果, 标记"待确认"
"""
import re
import os
from typing import List, Dict, Set, Optional

from utils.fuzzy_match import fuzzy_search, fuzzy_match_pair, normalize_text
from modules.project_normalizer import COUNTRIES


class FieldExtractor:
    def __init__(
        self,
        agent_email_map: Dict[str, dict],
        project_names: List[dict],
        logger=None,
        llm_client=None,
        ocr_fallback: bool = True,
    ):
        """
        agent_email_map: {邮箱地址(lower): {"代理": str, "代理简称": str, "收件人邮箱": str}}
        project_names: [{"项目编号": str, "项目名称": str, "国家": str, "业务类型": str}, ...]
        llm_client: 可选的 LLM 客户端, 用于补充提取
        ocr_fallback: 是否启用 OCR 图片兜底（常规源+LLM 之后仍缺字段才触发）
        """
        self.agent_email_map = agent_email_map
        self.project_names = project_names
        self.project_name_list = [p["项目名称"] for p in project_names if p.get("项目名称")]
        self.project_candidates = [(p["项目名称"], p) for p in project_names if p.get("项目名称")]
        self.logger = logger
        self.llm_client = llm_client
        self.ocr_fallback = ocr_fallback

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def extract_fields(self, mail: dict) -> List[Dict]:
        """
        从单封邮件提取字段
        返回: [{代理, 客户, 项目原始值, 项目(标准化), 需求, 代理匹配方式,
                客户提取来源, 置信度, match_candidates, ...}]
        """
        subject = mail.get("subject", "") or ""
        body = mail.get("body_text", "") or ""
        attachments = mail.get("attachments", [])
        sender_email = mail.get("sender_email", "").lower()

        # E1a: 提取需求类型
        need = self._extract_need(subject + " " + body)

        # E1b: 提取代理
        agent_result = self._extract_agent(sender_email, subject, body, attachments)

        # E1c: 提取客户
        customer_result = self._extract_customer(subject, body, attachments)

        # E1d: 提取项目
        project_result = self._extract_projects(subject, body, attachments)

        # E2: LLM 补充提取 — 仅对缺失字段
        llm_used = False
        missing_agent = not agent_result.get("agent") or agent_result.get("match_method") == "待确认"
        missing_customer = not customer_result.get("customer") or customer_result.get("source") == "待确认"
        missing_project = not project_result.get("projects")
        missing_need = not need

        if self.llm_client and self.llm_client.enabled and (missing_agent or missing_customer or missing_project or missing_need):
            self._log(f"规则提取有缺失, 调 LLM 补充: {subject[:40]}")
            llm_result = self.llm_client.extract_fields_llm(subject, body, attachments, sender_email)
            if llm_result:
                llm_used = True
                # 仅补充缺失字段, 不覆盖已有结果
                if missing_agent and llm_result.get("代理"):
                    agent_result["agent"] = llm_result["代理"]
                    agent_result["match_method"] = "LLM补充"
                if missing_customer and llm_result.get("客户"):
                    customer_result["customer"] = llm_result["客户"]
                    customer_result["source"] = "LLM补充"
                if missing_need and llm_result.get("需求"):
                    need = llm_result["需求"]
                if missing_project and llm_result.get("项目"):
                    llm_projects = llm_result["项目"]
                    project_result = self._match_llm_projects(llm_projects)

        # E3: OCR 图片兜底 — 常规源(标题/正文/文档附件/文件名)+LLM 全部读完仍缺字段才触发
        # 触发条件只看 客户/项目（OCR 能实际补的字段）; 代理缺失不单独触发
        # （代理主要靠邮箱表匹配, 表未命中时几乎每封都缺, 单独触发会使兜底退化为常态）
        # 代理在兜底过程中机会性提取（OCR 文本中的邮箱查代理表）
        ocr_used = False
        still_missing = (
            not customer_result.get("customer")
            or not project_result.get("projects")
        )
        if self.ocr_fallback and still_missing:
            ocr_used = self._ocr_fallback(attachments, agent_result, customer_result, project_result)

        # 组装输出
        rows = []
        if project_result["projects"]:
            for proj in project_result["projects"]:
                rows.append({
                    "sender_email": mail.get("sender_email", ""),
                    "sender_name": mail.get("sender_name", ""),
                    "date": mail.get("date"),
                    "subject": subject,
                    "body_text": body,
                    "代理": agent_result["agent"],
                    "代理匹配方式": agent_result["match_method"],
                    "客户": customer_result["customer"],
                    "客户提取来源": customer_result["source"],
                    "项目": proj["standard_name"],
                    "项目原始值": proj["raw_value"],
                    "需求": need,
                    "置信度": "low" if ocr_used else self._calc_confidence(agent_result, customer_result, proj),
                    "match_candidates": agent_result.get("candidates", []),
                    "llm_used": llm_used,
                    "ocr_used": ocr_used,
                    "uid": mail.get("uid"),
                    "filter_status": mail.get("filter_status"),
                    "filter_reason": mail.get("filter_reason", ""),
                })
        else:
            # 没提取到项目，也要输出一行
            rows.append({
                "sender_email": mail.get("sender_email", ""),
                "sender_name": mail.get("sender_name", ""),
                "date": mail.get("date"),
                "subject": subject,
                "body_text": body,
                "代理": agent_result["agent"],
                "代理匹配方式": agent_result["match_method"],
                "客户": customer_result["customer"],
                "客户提取来源": customer_result["source"],
                "项目": "",
                "项目原始值": "",
                "需求": need,
                "置信度": "low",
                "match_candidates": agent_result.get("candidates", []),
                "llm_used": llm_used,
                "ocr_used": ocr_used,
                "uid": mail.get("uid"),
                "filter_status": mail.get("filter_status"),
                "filter_reason": mail.get("filter_reason", ""),
            })

        return rows

    def _ocr_fallback(self, attachments: list, agent_result: dict,
                      customer_result: dict, project_result: dict) -> bool:
        """
        E3: OCR 图片兜底提取
        - 只在常规源+LLM 之后仍缺失字段时被调用
        - 只补缺失字段, 不覆盖已有结果
        - 结果标记来源=OCR图片兜底, 不作为可信真值（后续人工复核）
        - 一张图识别出多个项目 → project_result 收多条 → 上层拆多行输出
        - 全部图片识别后仍无目标信息 → 返回 False（保持原空结果）
        """
        pending = [a for a in attachments if a.get("ocr_pending") and a.get("filepath")]
        if not pending:
            return False

        self._log(f"字段缺失, 启动OCR兜底: {len(pending)} 张图片")
        filled = False
        for att in pending:
            text = self._ocr_image(att["filepath"])
            att["ocr_pending"] = False
            if text:
                att["text_content"] = text
            if not text:
                continue

            # 代理: OCR 文本中的邮箱 → 查代理表
            if not agent_result.get("agent"):
                try:
                    from utils.attachment_parser import extract_emails_from_text
                    for em in extract_emails_from_text(text):
                        entry = self.agent_email_map.get(em.lower())
                        if entry and entry.get("代理"):
                            agent_result["agent"] = entry["代理"]
                            agent_result["match_method"] = "OCR图片兜底"
                            filled = True
                            break
                except Exception:
                    pass

            # 客户: 公司名模式抽取
            if not customer_result.get("customer"):
                cust = self._extract_company_from_attachment_text(text)
                if cust:
                    customer_result["customer"] = cust
                    customer_result["source"] = "OCR图片兜底"
                    filled = True

            # 项目: 附件四标准名 + 离线规则(国家+业务), 多项目收集后由上层拆行
            if not project_result.get("projects"):
                projs = self._extract_projects_by_rules(text)
                existing = {p["standard_name"] for p in projs}
                for name in self.project_name_list:
                    if name and name in text and name not in existing:
                        projs.append({"raw_value": name, "standard_name": name})
                        existing.add(name)
                if projs:
                    project_result["projects"] = projs
                    filled = True

        if not filled:
            self._log("OCR兜底未提取到目标字段", "warning")
        return filled

    def _ocr_image(self, filepath: str) -> str:
        """调用 PaddleOCR 识别单张图片, 失败返回空串（不影响主流程）"""
        try:
            from utils.attachment_parser import _parse_image
            return _parse_image(filepath) or ""
        except Exception as e:
            self._log(f"OCR兜底识别失败 {filepath}: {e}", "warning")
            return ""

    def _match_llm_projects(self, llm_projects: List[str]) -> dict:
        """将 LLM 返回的项目名与附件四做匹配标准化"""
        matched = []
        for proj in llm_projects:
            proj_norm = normalize_text(proj)
            best_match = None
            best_score = 0
            for std_name in self.project_name_list:
                std_norm = normalize_text(std_name)
                if proj_norm == std_norm:
                    best_match = std_name
                    best_score = 100
                    break
                if proj_norm in std_norm or std_norm in proj_norm:
                    if len(std_name) > (best_score / 2):
                        best_match = std_name
                        best_score = len(std_name) * 2
            if best_match:
                matched.append({"raw_value": proj, "standard_name": best_match})
            else:
                matched.append({"raw_value": proj, "standard_name": proj})

        # 去重
        seen = set()
        unique = []
        for p in matched:
            if p["standard_name"] not in seen:
                seen.add(p["standard_name"])
                unique.append(p)
        return {"projects": unique}

    def _extract_need(self, text: str) -> str:
        """提取需求类型: 注册/新增/撤单"""
        if "撤单" in text:
            return "撤单"
        if "新增" in text:
            return "新增"
        if "注册" in text:
            return "注册"
        return ""

    def _extract_agent(
        self, sender_email: str, subject: str, body: str, attachments: list
    ) -> dict:
        """
        提取代理: 通过发件人邮箱查附件三
        逻辑: 精确匹配 → 模糊匹配 → 多匹配展示
        """
        def _get_agent_name(entry):
            if isinstance(entry, dict):
                return entry.get("代理", "") or entry.get("agent", "")
            return str(entry) if entry else ""

        # 精确匹配
        if sender_email in self.agent_email_map:
            entry = self.agent_email_map[sender_email]
            return {
                "agent": _get_agent_name(entry),
                "match_method": "邮箱精确匹配",
                "candidates": [],
            }

        # 也尝试从正文中提取邮箱地址
        from utils.attachment_parser import extract_emails_from_text
        body_emails = extract_emails_from_text(body)
        for em in body_emails:
            em_lower = em.lower()
            if em_lower in self.agent_email_map:
                entry = self.agent_email_map[em_lower]
                return {
                    "agent": _get_agent_name(entry),
                    "match_method": "正文邮箱精确匹配",
                    "candidates": [],
                }

        # 模糊匹配 — 用发件人邮箱去模糊匹配代理邮箱表
        candidates = []
        for em, entry in self.agent_email_map.items():
            score = self._email_similarity(sender_email, em)
            if score >= 80:
                candidates.append({
                    "text": em,
                    "data": entry,
                    "score": score,
                })

        candidates.sort(key=lambda x: x["score"], reverse=True)

        if not candidates:
            self._log(f"代理未匹配: {sender_email}", "warning")
            return {
                "agent": "",
                "match_method": "待确认",
                "candidates": [],
            }

        if len(candidates) == 1 and candidates[0]["score"] >= 90:
            entry = candidates[0]["data"]
            return {
                "agent": _get_agent_name(entry),
                "match_method": "邮箱模糊匹配",
                "candidates": [],
            }

        # 多匹配 — 全部展示
        self._log(f"代理多匹配({len(candidates)}条): {sender_email}", "warning")
        return {
            "agent": " / ".join(_get_agent_name(c["data"]) for c in candidates),
            "match_method": "多匹配-人工排查",
            "candidates": candidates,
        }

    def _email_similarity(self, a: str, b: str) -> int:
        """邮箱地址相似度比较"""
        from rapidfuzz import fuzz
        a_norm = normalize_text(a)
        b_norm = normalize_text(b)
        return int(fuzz.ratio(a_norm, b_norm))

    def _extract_customer(
        self, subject: str, body: str, attachments: list
    ) -> dict:
        """
        提取客户(公司名称)
        优先级: 主题 → 正文 → 附件(申请表/文件名)
        """
        # 从主题提取 — 主题格式通常是: 代理+客户号+公司名+项目+注册
        cust = self._extract_company_from_subject(subject)
        if cust:
            return {"customer": cust, "source": "主题"}

        # 从正文提取
        cust = self._extract_company_from_body(body)
        if cust:
            return {"customer": cust, "source": "正文"}

        # 从附件提取
        for att in attachments:
            # 从附件文件名提取
            fname = att.get("filename", "")
            cust = self._extract_company_from_filename(fname)
            if cust:
                return {"customer": cust, "source": f"附件文件名({fname})"}

            # 从附件内容提取
            text = att.get("text_content", "")
            cust = self._extract_company_from_attachment_text(text)
            if cust:
                return {"customer": cust, "source": "附件内容"}

        return {"customer": "", "source": "待确认"}

    def _extract_company_from_subject(self, subject: str) -> str:
        """从主题中提取公司名"""
        if not subject:
            return ""

        # 尝试按常见分隔符切分
        parts = re.split(r"[+＋\-\|]", subject)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            # 公司名特征: 含"公司"或"有限"或"科技"等
            if any(kw in part for kw in ["公司", "有限", "科技", "电商", "贸易", "实业"]):
                # 排除代理名
                if not self._is_agent_name(part):
                    return part
        return ""

    def _extract_company_from_body(self, body: str) -> str:
        """从正文中提取公司名"""
        if not body:
            return ""
        # 搜索含公司名特征的行
        lines = body.split("\n")
        for line in lines:
            line = line.strip()
            if any(kw in line for kw in ["公司", "有限", "科技", "电商", "贸易", "实业"]):
                # 排除代理名和邮件签名
                if not self._is_agent_name(line) and len(line) < 50:
                    # 清理行内多余内容
                    match = re.search(r"([\u4e00-\u9fa5A-Za-z]+(?:公司|有限[\u4e00-\u9fa5]*公司))", line)
                    if match:
                        return match.group(1)
        return ""

    def _extract_company_from_filename(self, filename: str) -> str:
        """从附件文件名中提取公司名"""
        if not filename:
            return ""
        # 去除扩展名
        name = os.path.splitext(filename)[0] if "." in filename else filename
        if any(kw in name for kw in ["公司", "有限", "科技", "电商", "贸易", "实业"]):
            match = re.search(r"([\u4e00-\u9fa5A-Za-z]+(?:公司|有限[\u4e00-\u9fa5]*公司))", name)
            if match:
                return match.group(1)
        return ""

    def _extract_company_from_attachment_text(self, text: str) -> str:
        """从附件文本中提取公司名"""
        if not text:
            return ""
        # 搜索"公司名称"标签后的内容
        patterns = [
            r"公司名称[:\s：]*([^\n\r，,]{2,30})",
            r"甲方[:\s：]*([^\n\r，,]{2,30})",
            r"申请单位[:\s：]*([^\n\r，,]{2,30})",
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                name = match.group(1).strip()
                if any(kw in name for kw in ["公司", "有限", "科技", "电商", "贸易"]):
                    return name
        # 兜底: 无标签的裸公司名（OCR 截图常见, 如订单确认截图里直接印公司名）
        # 取最长匹配（避免截到公司名的前缀子串）
        bare = re.findall(
            r"[\u4e00-\u9fa5A-Za-z0-9（）()]{2,24}(?:(?:有限|股份)责任?|责任)?公司", text
        )
        if bare:
            return max(bare, key=len)
        return ""

    def _is_agent_name(self, text: str) -> bool:
        """判断文本是否是代理名(而非公司名)"""
        agent_names = set()
        for entry in self.agent_email_map.values():
            if isinstance(entry, dict):
                name = entry.get("代理", "") or entry.get("agent", "")
                if name:
                    agent_names.add(name)
                short = entry.get("代理简称", "")
                if short:
                    agent_names.add(short)
            elif isinstance(entry, str):
                agent_names.add(entry)
        return text.strip() in agent_names

    def _extract_projects(self, subject: str, body: str, attachments: list) -> dict:
        """
        提取项目: 用附件四项目名列表扫描主题/正文/附件
        """
        search_text = subject + " " + body
        for att in attachments:
            search_text += " " + att.get("text_content", "")
            search_text += " " + att.get("filename", "")

        matched_projects = []

        for proj_name in self.project_name_list:
            if not proj_name:
                continue
            if proj_name in search_text:
                matched_projects.append({
                    "raw_value": proj_name,
                    "standard_name": proj_name,
                })

        # 去重
        seen = set()
        unique = []
        for p in matched_projects:
            key = p["standard_name"]
            if key not in seen:
                seen.add(key)
                unique.append(p)

        # 附件四未命中（或未导入附件四）→ 离线规则提取: "国家+业务类型" 相邻模式
        if not unique:
            unique = self._extract_projects_by_rules(search_text)
            if unique:
                self._log(f"规则提取项目(无附件四匹配): {[p['standard_name'] for p in unique]}")

        if not unique:
            self._log(f"项目未匹配: {subject[:40]}", "warning")

        return {"projects": unique}

    # 业务类型别名表: 标准名 → 正则别名
    _BT_PATTERNS = [
        ("WEEE", r"weee"),
        ("电池法", r"电池法|电池"),
        ("包装法", r"包装法|包装"),
        ("一次性塑料", r"一次性塑料(?:法)?|塑料法"),
        ("EPR", r"epr"),
    ]

    def _extract_projects_by_rules(self, text: str) -> List[Dict]:
        """
        离线规则提取项目: 不依赖附件四，扫描 "国家+业务类型" 相邻模式
        覆盖: "德国WEEE"、"法国包装法"、"波兰荷兰包装法"(国家连排)、"WEEE德国"(反序)
        """
        if not text:
            return []
        text_l = text.lower()
        country_alt = "|".join(COUNTRIES)
        found = []
        seen = set()

        def _add(raw: str, std: str):
            if std and std not in seen:
                seen.add(std)
                found.append({"raw_value": raw, "standard_name": std})

        for std_bt, bt_pat in self._BT_PATTERNS:
            # 模式1: 国家+业务 (含国家连排: "波兰荷兰包装法")
            for m in re.finditer(f"((?:{country_alt})+)(?:{bt_pat})", text_l):
                seg = m.group(1)
                for c in COUNTRIES:
                    if c in seg:
                        _add(m.group(0), c + std_bt)
            # 模式2: 业务+国家 (反序, 如 "WEEE德国")
            for m in re.finditer(f"(?:{bt_pat})((?:{country_alt})+)", text_l):
                seg = m.group(1)
                for c in COUNTRIES:
                    if c in seg:
                        _add(m.group(0), c + std_bt)

        # 模式3: 国家+业务1(+业务2...) 共享国家前缀 (如 "荷兰WEEE+包装法"、"德国WEEE/电池法/包装法")
        all_bt = "|".join(pat for _, pat in self._BT_PATTERNS)
        alias_to_std = {}
        for std_bt, pat in self._BT_PATTERNS:
            for alias in pat.split("|"):
                alias_to_std[alias] = std_bt
        for m in re.finditer(f"({country_alt})((?:{all_bt})(?:\\s*[+＋/／、，,]\\s*(?:{all_bt}))+)", text_l):
            c = m.group(1)
            for bt_m in re.finditer(all_bt, m.group(2)):
                _add(m.group(0), c + alias_to_std[bt_m.group(0)])

        return found

    def _calc_confidence(self, agent_r: dict, customer_r: dict, proj_r: dict) -> str:
        """计算置信度"""
        score = 0
        if agent_r.get("agent"):
            score += 1
        if customer_r.get("customer"):
            score += 1
        if proj_r.get("standard_name"):
            score += 1

        if agent_r.get("match_method", "").startswith("多匹配"):
            return "low"
        if score == 3:
            return "high"
        if score >= 2:
            return "medium"
        return "low"
