"""M1 邮件读取模块 — 连接阿里企业邮箱 IMAP，拉取收件箱邮件"""
import os
import email
import imaplib
import tempfile
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
from typing import List, Dict

from bs4 import BeautifulSoup


def _decode_str(s):
    """解码邮件头部字符串"""
    if s is None:
        return ""
    parts = decode_header(s)
    result = []
    for data, charset in parts:
        if isinstance(data, bytes):
            try:
                result.append(data.decode(charset or "utf-8", errors="replace"))
            except (LookupError, Exception):
                result.append(data.decode("utf-8", errors="replace"))
        else:
            result.append(data)
    return "".join(result)


def _html_to_text(html_str):
    """HTML转纯文本"""
    soup = BeautifulSoup(html_str, "lxml")
    for tag in soup.find_all(["style", "script"]):
        tag.decompose()
    return soup.get_text(separator="\n", strip=True)


def _simplify_body(body_text):
    """精简正文 — 去除常见签名/寒暄"""
    lines = body_text.split("\n")
    kept = []
    skip_patterns = [
        "请审核资料",
        "有问题请联系",
        "谢谢",
        "祝好",
        "best regards",
        "regards",
        "--",
        "发件人",
        "发件时间",
        "收件人",
        "主题",
        "本邮件及其附件",
        "confidential",
    ]
    for line in lines:
        line = line.strip()
        if not line:
            continue
        lower = line.lower()
        if any(p in lower for p in skip_patterns):
            continue
        kept.append(line)
    return "\n".join(kept) if kept else body_text[:500]


class MailReader:
    def __init__(self, config: dict, logger=None):
        self.server_addr = config["imap_server"]
        self.port = config["imap_port"]
        self.address = config["address"]
        self.password = config["password"]
        self.mailbox = config.get("mailbox", "INBOX")
        self.logger = logger

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def fetch_mails(
        self,
        date_from: datetime,
        date_to: datetime,
        self_email: str,
        progress_callback=None,
    ) -> List[Dict]:
        """
        拉取指定时间范围内的收件箱邮件
        date_from / date_to: 时间范围
        self_email: 自身邮箱地址，用于跳过自身发送的邮件
        progress_callback: 可选回调(current, total)
        返回: List[mail_dict]
        """
        self._log(f"连接邮箱 {self.server_addr}:{self.port} ...")
        conn = None
        try:
            conn = imaplib.IMAP4_SSL(self.server_addr, self.port)
            conn.login(self.address, self.password)
            self._log("邮箱连接成功")
        except Exception as e:
            self._log(f"邮箱连接失败: {e}", "error")
            raise

        try:
            conn.select(self.mailbox, readonly=True)
            search_from = date_from.strftime("%d-%b-%Y")
            search_to = (date_to + timedelta(days=1)).strftime("%d-%b-%Y")
            status, data = conn.search(None, f'SINCE {search_from}', f'BEFORE {search_to}')

            if status != "OK":
                self._log("搜索邮件失败", "error")
                return []

            uids = data[0].split()
            total = len(uids)
            self._log(f"找到 {total} 封邮件 ({date_from.date()} ~ {date_to.date()})")

            mails = []
            for idx, uid_bytes in enumerate(uids):
                uid = uid_bytes.decode()
                status, msg_data = conn.fetch(uid_bytes, "(RFC822)")
                if status != "OK" or not msg_data or not msg_data[0]:
                    continue

                raw = msg_data[0][1]
                msg = email.message_from_bytes(raw)

                sender = _decode_str(msg.get("From", ""))
                sender_email = self._extract_email_addr(sender)

                # 跳过自身发送的邮件
                if self_email.lower() in sender_email.lower():
                    mails.append({
                        "uid": uid,
                        "sender_email": sender_email,
                        "sender_name": sender,
                        "date": parsedate_to_datetime(msg.get("Date", "")),
                        "subject": _decode_str(msg.get("Subject", "")),
                        "body_text": "",
                        "body_raw": "",
                        "attachments": [],
                        "skip_reason": "self_sent",
                    })
                    if progress_callback:
                        progress_callback(idx + 1, total)
                    continue

                subject = _decode_str(msg.get("Subject", ""))
                date = parsedate_to_datetime(msg.get("Date", ""))
                body_text, body_raw, attachments = self._parse_msg_content(msg)

                mails.append({
                    "uid": uid,
                    "sender_email": sender_email,
                    "sender_name": sender,
                    "date": date,
                    "subject": subject,
                    "body_text": body_text,
                    "body_raw": body_raw,
                    "attachments": attachments,
                    "skip_reason": None,
                })

                if progress_callback:
                    progress_callback(idx + 1, total)

            self._log(f"邮件解析完成, 共 {len(mails)} 封 (含跳过)")
            return mails
        finally:
            try:
                conn.close()
            except Exception:
                pass
            try:
                conn.logout()
            except Exception:
                pass

    def _extract_email_addr(self, sender_str: str) -> str:
        """从发件人字段提取邮箱地址"""
        import re
        match = re.search(r"<([^>]+)>", sender_str)
        if match:
            return match.group(1).strip()
        match = re.search(r"[\w.+-]+@[\w.-]+\.\w+", sender_str)
        if match:
            return match.group(0).strip()
        return sender_str.strip().lower()

    def _parse_msg_content(self, msg) -> tuple:
        """解析邮件正文和附件"""
        body_text = ""
        body_raw = ""
        attachments = []

        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = str(part.get("Content-Disposition", ""))
                content_type = part.get_content_type()
                filename = part.get_filename()
                filename = _decode_str(filename) if filename else None

                if "attachment" in content_disposition.lower() or filename:
                    if filename:
                        payload = part.get_payload(decode=True)
                        if payload:
                            with tempfile.NamedTemporaryFile(
                                delete=False, suffix=os.path.splitext(filename)[1]
                            ) as tmp:
                                tmp.write(payload)
                                tmp_path = tmp.name
                            try:
                                from utils.attachment_parser import parse_attachment
                                att = parse_attachment(tmp_path, filename)
                                att["filename"] = filename
                                attachments.append(att)
                            finally:
                                try:
                                    os.unlink(tmp_path)
                                except OSError:
                                    pass
                elif content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload and not body_text:
                        charset = part.get_content_charset() or "utf-8"
                        body_raw = payload.decode(charset, errors="replace")
                        body_text = body_raw
                elif content_type == "text/html" and not body_text:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        html_str = payload.decode(charset, errors="replace")
                        body_raw = html_str
                        body_text = _html_to_text(html_str)
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                content_type = msg.get_content_type()
                raw = payload.decode(charset, errors="replace")
                body_raw = raw
                if content_type == "text/html":
                    body_text = _html_to_text(raw)
                else:
                    body_text = raw

        body_text = _simplify_body(body_text)
        return body_text, body_raw, attachments
