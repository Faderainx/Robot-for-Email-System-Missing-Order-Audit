"""M3 字段提取模块 — 表驱动匹配，从邮件主题/正文/附件提取关键字段"""
import re
from typing import List, Dict, Set, Optional

from utils.fuzzy_match import fuzzy_search, fuzzy_match_pair, normalize_text


class FieldExtractor:
    def __init__(
        self,
        agent_email_map: Dict[str, dict],
        project_names: List[dict],
        logger=None,
    ):
        """
        agent_email_map: {邮箱地址(lower): {"代理": str, "代理简称": str, "收件人邮箱": str}}
        project_names: [{"项目编号": str, "项目名称": str, "国家": str, "业务类型": str}, ...]
        """
        self.agent_email_map = agent_email_map
        self.project_names = project_names
        self.project_name_list = [p["项目名称"] for p in project_names if p.get("项目名称")]
        self.project_candidates = [(p["项目名称"], p) for p in project_names if p.get("项目名称")]
        self.logger = logger

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

        # 提取需求类型
        need = self._extract_need(subject + " " + body)

        # 提取代理
        agent_result = self._extract_agent(sender_email, subject, body, attachments)

        # 提取客户
        customer_result = self._extract_customer(subject, body, attachments)

        # 提取项目
        project_result = self._extract_projects(subject, body, attachments)

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
                    "置信度": self._calc_confidence(agent_result, customer_result, proj),
                    "match_candidates": agent_result.get("candidates", []),
                    "uid": mail.get("uid"),
                    "filter_status": mail.get("filter_status"),
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
                "uid": mail.get("uid"),
                "filter_status": mail.get("filter_status"),
            })

        return rows

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
        # 精确匹配
        if sender_email in self.agent_email_map:
            entry = self.agent_email_map[sender_email]
            return {
                "agent": entry.get("代理", ""),
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
                    "agent": entry.get("代理", ""),
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
                "agent": entry.get("代理", ""),
                "match_method": "邮箱模糊匹配",
                "candidates": [],
            }

        # 多匹配 — 全部展示
        self._log(f"代理多匹配({len(candidates)}条): {sender_email}", "warning")
        return {
            "agent": " / ".join(c["data"].get("代理", "") for c in candidates),
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
        return ""

    def _is_agent_name(self, text: str) -> bool:
        """判断文本是否是代理名(而非公司名)"""
        agent_names = set()
        for entry in self.agent_email_map.values():
            name = entry.get("代理", "")
            if name:
                agent_names.add(name)
                short = entry.get("代理简称", "")
                if short:
                    agent_names.add(short)
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

        if not unique:
            self._log(f"项目未匹配: {subject[:40]}", "warning")

        return {"projects": unique}

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


import os
