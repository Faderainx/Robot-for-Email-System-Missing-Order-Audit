"""M2 邮件过滤模块 — 识别注册/新增/撤单类有效邮件，排除无关邮件"""
from typing import List, Dict, Set

# 方法一：主题/正文含注册类关键字
KEYWORDS_METHOD1 = {
    "注册", "新增", "撤单", "新增品牌", "新增类别",
}

# 方法二：主题/正文/附件含业务关键字
KEYWORDS_METHOD2 = {
    "客户号", "代理", "客户", "项目",
    "WEEE", "电池法", "包装法", "EPR",
    "一次性塑料", "担保", "回收费",
    "weee", "epr", "battery", "packaging",
}

# 外部无关邮件特征词
IRRELEVANT_KEYWORDS = {
    "账单", "invoice", "合同", "consulting",
    "业务咨询", "反馈", "下证",
}


class MailFilter:
    def __init__(self, internal_emails: Set[str], logger=None):
        """
        internal_emails: 公司内部邮箱地址集合 (附件二)
        """
        self.internal_emails = {e.lower().strip() for e in internal_emails}
        self.logger = logger

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def filter_mails(self, mails: List[Dict]) -> tuple:
        """
        过滤邮件
        返回: (valid_mails, filtered_mails)
        valid_mails: 有效邮件 (含不确定的)
        filtered_mails: 被过滤邮件
        """
        valid_mails = []
        filtered_mails = []

        for mail in mails:
            # 跳过自身发送的邮件
            if mail.get("skip_reason") == "self_sent":
                filtered_mails.append({
                    **mail,
                    "filter_reason": "自身发送",
                })
                continue

            subject = mail.get("subject", "") or ""
            body = mail.get("body_text", "") or ""
            attachments = mail.get("attachments", [])

            att_text = " ".join(a.get("text_content", "") for a in attachments)
            all_text = subject + " " + body + " " + att_text

            # F1: 方法一 — 主题/正文含注册类关键字
            if self._contains_any(subject + " " + body, KEYWORDS_METHOD1):
                mail["filter_status"] = "valid"
                mail["filter_reason"] = "方法一: 注册/新增/撤单关键字命中"
                valid_mails.append(mail)
                self._log(f"有效(方法一): {subject[:40]}")
                continue

            # F2: 方法二 — 主题/正文/附件含业务关键字
            if self._contains_any(all_text, KEYWORDS_METHOD2):
                mail["filter_status"] = "valid"
                mail["filter_reason"] = "方法二: 业务关键字命中"
                valid_mails.append(mail)
                self._log(f"有效(方法二): {subject[:40]}")
                continue

            # F3: 内部无关邮件排除
            sender = mail.get("sender_email", "").lower()
            if sender in self.internal_emails:
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "内部无关邮件"
                filtered_mails.append(mail)
                self._log(f"过滤(内部无关): {subject[:40]}")
                continue

            # F4: 外部无关邮件排除
            if self._contains_any(all_text, IRRELEVANT_KEYWORDS) and not self._contains_any(
                all_text, {"代理", "客户", "项目", "WEEE", "电池法", "包装法", "EPR"}
            ):
                mail["filter_status"] = "filtered"
                mail["filter_reason"] = "外部无关邮件"
                filtered_mails.append(mail)
                self._log(f"过滤(外部无关): {subject[:40]}")
                continue

            # F5: 不确定 — 保留待人工复核
            mail["filter_status"] = "uncertain"
            mail["filter_reason"] = "不确定, 需人工复核"
            valid_mails.append(mail)
            self._log(f"不确定(保留): {subject[:40]}", "warning")

        self._log(f"过滤完成: 有效 {len(valid_mails)} 封, 过滤 {len(filtered_mails)} 封")
        return valid_mails, filtered_mails

    def _contains_any(self, text: str, keywords) -> bool:
        text_lower = text.lower()
        for kw in keywords:
            if kw.lower() in text_lower:
                return True
        return False
