"""M1 邮件读取模块 — 连接阿里企业邮箱 IMAP，拉取收件箱邮件

关键设计: 先批量拉取全部邮件原始数据并断开连接，再离线解析附件
避免因附件解析耗时过长导致 IMAP 连接超时断开
"""
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
        # 原始邮件本地缓存（按 mailbox+uid 落盘，重跑同范围直接命中，断连不丢已拉数据）
        self.cache_dir = config.get("cache_dir") or os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "cache", "mails"
        )
        # 每拉取 N 封（真实网络拉取，不含缓存命中）重连一次
        self.reconnect_batch_size = config.get("reconnect_batch_size", 20)
        self.logger = logger

    def _log(self, msg, level="info"):
        if self.logger:
            getattr(self.logger, level)(msg)

    def _connect(self, ctx):
        """创建新的 IMAP 连接"""
        conn = imaplib.IMAP4_SSL(self.server_addr, self.port, ssl_context=ctx, timeout=30)
        conn.login(self.address, self.password)
        conn.select(self.mailbox, readonly=True)
        return conn

    def fetch_mails(
        self,
        date_from,
        date_to,
        self_email: str,
        progress_callback=None,
    ) -> List[Dict]:
        """
        拉取指定时间范围内的收件箱邮件
        分批拉取: 每 20 封重连一次，避免连接超时
        date_from/date_to: 支持 datetime 对象或 'YYYY-MM-DD' 字符串
        """
        if isinstance(date_from, str):
            date_from = datetime.strptime(date_from, "%Y-%m-%d")
        if isinstance(date_to, str):
            date_to = datetime.strptime(date_to, "%Y-%m-%d")

        import ssl
        ctx = ssl.create_default_context()
        conn = None
        raw_mails = []

        try:
            conn = self._connect(ctx)
            self._log("邮箱连接成功")

            search_from = date_from.strftime("%d-%b-%Y")
            search_to = (date_to + timedelta(days=1)).strftime("%d-%b-%Y")
            status, data = conn.search(None, f'SINCE {search_from}', f'BEFORE {search_to}')

            if status != "OK":
                self._log("搜索邮件失败", "error")
                return []

            uids = data[0].split()
            total = len(uids)
            self._log(f"找到 {total} 封邮件 ({date_from.date()} ~ {date_to.date()})")

            os.makedirs(self.cache_dir, exist_ok=True)
            cache_hits = 0
            fetched = 0  # 真实网络拉取计数（重连阈值只看这个，缓存命中不触发重连）
            BATCH_SIZE = self.reconnect_batch_size

            for idx, uid_bytes in enumerate(uids):
                uid = uid_bytes.decode()

                # 缓存命中 → 直接读本地，跳过网络
                cache_path = os.path.join(self.cache_dir, f"{self.mailbox}_{uid}.eml")
                if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
                    try:
                        with open(cache_path, "rb") as f:
                            raw_mails.append((uid, f.read()))
                        cache_hits += 1
                    except OSError as e:
                        self._log(f"读缓存失败(uid={uid}): {e}", "warning")
                    if progress_callback:
                        progress_callback(idx + 1, total)
                    continue

                # 每拉取 BATCH_SIZE 封重连一次
                if fetched > 0 and fetched % BATCH_SIZE == 0:
                    try:
                        conn.close()
                    except Exception:
                        pass
                    try:
                        conn.logout()
                    except Exception:
                        pass
                    self._log(f"重连邮箱 (已拉取 {idx}/{total})...")
                    conn = self._connect(ctx)

                try:
                    status, msg_data = conn.fetch(uid_bytes, "(RFC822)")
                except Exception as e:
                    self._log(f"拉取第{idx+1}封失败(uid={uid}): {e}", "warning")
                    # 尝试重连后重试一次
                    try:
                        conn = self._connect(ctx)
                        status, msg_data = conn.fetch(uid_bytes, "(RFC822)")
                    except Exception as e2:
                        self._log(f"重试失败: {e2}", "warning")
                        continue
                if status != "OK" or not msg_data or not msg_data[0]:
                    continue

                raw_bytes = msg_data[0][1]
                raw_mails.append((uid, raw_bytes))
                fetched += 1

                # 落盘缓存（失败不影响主流程）
                try:
                    with open(cache_path, "wb") as f:
                        f.write(raw_bytes)
                except OSError as e:
                    self._log(f"写缓存失败(uid={uid}): {e}", "warning")

                if progress_callback:
                    progress_callback(idx + 1, total)

            self._log(
                f"原始邮件拉取完成: {len(raw_mails)} 封 "
                f"(网络拉取 {fetched}, 缓存命中 {cache_hits})"
            )

        except Exception as e:
            self._log(f"邮箱连接失败: {e}", "error")
            raise
        finally:
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass
                try:
                    conn.logout()
                except Exception:
                    pass
            self._log("已断开邮箱连接, 开始离线解析...")

        # === 阶段二: 离线解析每封邮件的正文和附件 ===
        mails = []
        for idx, (uid, raw_bytes) in enumerate(raw_mails):
            msg = email.message_from_bytes(raw_bytes)

            sender = _decode_str(msg.get("From", ""))
            sender_email = self._extract_email_addr(sender)
            subject = _decode_str(msg.get("Subject", ""))
            date = parsedate_to_datetime(msg.get("Date", ""))

            # 跳过自身发送的邮件
            if self_email.lower() in sender_email.lower():
                mails.append({
                    "uid": uid,
                    "sender_email": sender_email,
                    "sender_name": sender,
                    "date": date,
                    "subject": subject,
                    "body_text": "",
                    "body_raw": "",
                    "attachments": [],
                    "skip_reason": "self_sent",
                })
                continue

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

            # 阶段二不再发进度回调（阶段一已按 total 汇报过，避免进度条重复回跳）
            if progress_callback:
                progress_callback(len(raw_mails), len(raw_mails))

        self._log(f"邮件解析完成, 共 {len(mails)} 封 (含跳过)")
        return mails

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
                            ext = os.path.splitext(filename)[1].lower()
                            if ext in (".jpg", ".jpeg", ".png", ".bmp", ".tiff"):
                                # 图片附件: 持久化到 cache/attachments/, 登记 ocr_pending
                                # OCR 仅在字段提取兜底阶段按需触发（懒加载）
                                from utils.attachment_parser import (
                                    parse_attachment, save_image_attachment,
                                )
                                img_path = save_image_attachment(payload, filename)
                                att = parse_attachment(img_path, filename)
                                att["filename"] = filename
                                attachments.append(att)
                            else:
                                with tempfile.NamedTemporaryFile(
                                    delete=False, suffix=ext
                                ) as tmp:
                                    tmp.write(payload)
                                    tmp_path = tmp.name
                                try:
                                    from utils.attachment_parser import parse_attachment
                                    att = parse_attachment(tmp_path, filename)
                                    att["filename"] = filename
                                    attachments.append(att)
                                    # 压缩包内的待识别图片合并进附件列表（懒加载 OCR）
                                    for p in att.pop("pending_images", []):
                                        attachments.append(p)
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
