"""M2 邮件过滤模块 — 规则引擎 + LLM 二次意图识别

流程:
  F1 规则过滤 → 产出 valid_mails + filtered_mails
  F2 LLM 二次识别 → 对 filtered_mails 逐封做意图分类
"""
import re
from typing import List, Dict, Set, Optional

# 方法一：主题含注册类关键字（必须同时含业务词才算有效）
KEYWORDS_ACTION = {"注册", "新增", "撤单", "注销", "追加"}

# 业务关键字（项目/法案名称）
KEYWORDS_BUSINESS = {
    "WEEE", "weee", "电池法", "电池", "包装法", "EPR", "epr",
    "一次性塑料", "BAT", "bat",
    "德国", "法国", "意大利", "西班牙", "荷兰", "波兰", "瑞典",
    "比利时", "爱尔兰", "葡萄牙", "奥地利",
}

# 明确无关的关键词（出现这些且不含业务词 → 过滤）
KEYWORDS_IRRELEVANT = {
    "账单", "invoice", "合同", "协议",
    "咨询", "反馈", "建议",
    "下证", "证书号", "证书下发",
    "保证金", "担保", "退款",
    "回收费", "购买",
}


class MailFilter:
    def __init__(self, internal_emails: Set[str], logger=None):
        self.internal_emails = {e.lower().strip() for e in internal_emails}
        self.logger = logger

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    @staticmethod
    def _contains_any(text: str, keywords) -> bool:
        text_lower = text.lower()
        for kw in keywords:
            if kw.lower() in text_lower:
                return True
        return False

    def filter_mails(self, mails: List[Dict]) -> tuple:
        valid_mails = []
        filtered_mails = []

        for mail in mails:
            if mail.get("skip_reason") == "self_sent":
                filtered_mails.append({**mail, "filter_reason": "自身发送"})
                continue

            subject = mail.get("subject", "") or ""
            body = mail.get("body_text", "") or ""
            attachments = mail.get("attachments", [])
            att_text = " ".join(a.get("text_content", "") for a in attachments)
            subject_body = subject + " " + body
            all_text = subject_body + " " + att_text

            sender = mail.get("sender_email", "").lower()

            # --- 判断逻辑 ---

            # 1. 无关邮件排除（含无关词且不含业务词）
            has_business = self._contains_any(all_text, KEYWORDS_BUSINESS)
            has_irrelevant = self._contains_any(all_text, KEYWORDS_IRRELEVANT)
            has_action = self._contains_any(subject_body, KEYWORDS_ACTION)

            if has_irrelevant and not has_business and not has_action:
                if sender in self.internal_emails:
                    reason = "内部无关邮件"
                else:
                    reason = "外部无关邮件"
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = reason
                filtered_mails.append(mail)
                self._log(f"过滤({reason}): {subject[:50]}")
                continue

            # 2. 方法一：主题/正文含注册类关键字 + 含业务词
            if has_action and has_business:
                mail["filter_status"] = "valid"
                mail["filter_reason"] = "方法一: 注册关键字+业务词命中"
                valid_mails.append(mail)
                self._log(f"有效(方法一): {subject[:50]}")
                continue

            # 3. 方法一扩展：主题含注册关键字（即使不含业务词，主题够明确）
            if has_action and self._contains_any(subject, KEYWORDS_ACTION):
                mail["filter_status"] = "valid"
                mail["filter_reason"] = "方法一扩展: 主题含注册关键字"
                valid_mails.append(mail)
                self._log(f"有效(方法一扩展): {subject[:50]}")
                continue

            # 4. 方法二：主题或正文含业务关键字（WEEE/电池法/包装法等）
            if has_business:
                # 但如果同时含无关词且无关词更突出 → 过滤
                if has_irrelevant and not has_action:
                    reason = "内部无关邮件" if sender in self.internal_emails else "外部无关邮件"
                    mail["filter_status"] = "filtered"
                    mail["filter_reason"] = f"{reason}(业务词+无关词混存)"
                    filtered_mails.append(mail)
                    self._log(f"过滤({reason}): {subject[:50]}")
                    continue
                mail["filter_status"] = "valid"
                mail["filter_reason"] = "方法二: 业务关键字命中"
                valid_mails.append(mail)
                self._log(f"有效(方法二): {subject[:50]}")
                continue

            # 5. 内部无关邮件
            if sender in self.internal_emails and not has_business:
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "内部无关邮件"
                filtered_mails.append(mail)
                self._log(f"过滤(内部无关): {subject[:50]}")
                continue

            # 6. 外部无关邮件
            if not has_business and not has_action:
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "外部无关邮件"
                filtered_mails.append(mail)
                self._log(f"过滤(外部无关): {subject[:50]}")
                continue

            # 7. 不确定 — 保留
            mail["filter_status"] = "uncertain"
            mail["filter_reason"] = "不确定, 需人工复核"
            valid_mails.append(mail)
            self._log(f"不确定(保留): {subject[:50]}", "warning")

        self._log(f"规则过滤完成: 有效 {len(valid_mails)} 封, 过滤 {len(filtered_mails)} 封")
        return valid_mails, filtered_mails

    def llm_second_pass(
        self,
        filtered_mails: List[Dict],
        llm_client,
        progress_callback=None,
    ) -> tuple:
        if not llm_client or not llm_client.enabled:
            self._log("LLM 未启用, 跳过二次识别", "warning")
            return [], filtered_mails

        if not filtered_mails:
            return [], []

        to_classify = [
            m for m in filtered_mails
            if m.get("filter_reason") != "自身发送"
        ]
        self._log(f"LLM 二次识别: 共 {len(to_classify)} 封待分类邮件")

        recovered = []
        still_filtered = [
            m for m in filtered_mails
            if m.get("filter_reason") == "自身发送"
        ]

        total = len(to_classify)
        for idx, mail in enumerate(to_classify, 1):
            subject = mail.get("subject", "") or ""
            body = mail.get("body_text", "") or ""
            attachments = mail.get("attachments", [])

            self._log(f"LLM 分类 {idx}/{total}: {subject[:50]}")
            result = llm_client.classify_email(subject, body, attachments)

            if result is None:
                mail["filter_reason"] = f"{mail.get('filter_reason','')} | LLM失败"
                still_filtered.append(mail)
            elif result.get("is_target"):
                intent = result.get("intent", "")
                reason = result.get("reason", "")
                mail["filter_status"] = "valid"
                mail["filter_reason"] = f"LLM恢复(意图={intent}, 原因={reason})"
                recovered.append(mail)
                self._log(f"LLM 恢复: {subject[:40]} → {intent}")
            else:
                email_type = result.get("email_type", "")
                reason = result.get("reason", "")
                mail["filter_reason"] = f"{mail.get('filter_reason','')} | LLM确认={email_type}({reason})"
                still_filtered.append(mail)
                self._log(f"LLM 确认过滤: {subject[:40]} → {email_type}")

            if progress_callback:
                progress_callback(idx, total)

        self._log(f"LLM 二次识别完成: 恢复 {len(recovered)} 封, 确认过滤 {len(still_filtered)} 封")
        return recovered, still_filtered
