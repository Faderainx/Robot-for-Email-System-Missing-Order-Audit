"""附件解析工具 — 支持 .zip .rar .xlsx .pdf .docx"""
import os
import tempfile
import zipfile
from typing import List, Dict

logger = None


def _get_logger():
    global logger
    if logger is None:
        import logging
        logger = logging.getLogger("mail_audit")
    return logger


def parse_attachment(filepath: str, filename: str) -> Dict:
    """
    解析附件文件，提取文本内容
    返回: {'filename': str, 'text_content': str, 'sheets': list}
    """
    result = {"filename": filename, "text_content": "", "sheets": []}
    if not os.path.exists(filepath):
        return result

    ext = os.path.splitext(filename)[1].lower()

    try:
        if ext == ".xlsx" or ext == ".xls":
            result.update(_parse_excel(filepath))
        elif ext == ".pdf":
            result["text_content"] = _parse_pdf(filepath)
        elif ext == ".docx":
            result["text_content"] = _parse_docx(filepath)
        elif ext == ".zip":
            result.update(_parse_zip(filepath))
        elif ext == ".rar":
            result.update(_parse_rar(filepath))
        elif ext == ".csv":
            result["text_content"] = _parse_csv(filepath)
        elif ext in (".txt",):
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                result["text_content"] = f.read()
        else:
            result["text_content"] = ""
            _get_logger().warning(f"不支持的附件格式: {filename}")
    except Exception as e:
        _get_logger().error(f"解析附件失败 {filename}: {e}")
        result["text_content"] = ""

    return result


def _parse_excel(filepath: str) -> Dict:
    from openpyxl import load_workbook
    text_parts = []
    sheets = []
    try:
        wb = load_workbook(filepath, read_only=True, data_only=True)
        for ws_name in wb.sheetnames:
            ws = wb[ws_name]
            rows_text = []
            for row in ws.iter_rows(max_row=200, values_only=True):
                cells = [str(c).strip() if c is not None else "" for c in row]
                line = " ".join(cells)
                if line.strip():
                    rows_text.append(line)
            sheet_text = "\n".join(rows_text)
            text_parts.append(sheet_text)
            sheets.append({"sheet_name": ws_name, "text": sheet_text})
        wb.close()
    except Exception as e:
        _get_logger().error(f"解析Excel失败: {e}")
    return {"text_content": "\n".join(text_parts), "sheets": sheets}


def _parse_pdf(filepath: str) -> str:
    import pdfplumber
    text_parts = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages[:30]:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)
    return "\n".join(text_parts)


def _parse_docx(filepath: str) -> str:
    from docx import Document
    doc = Document(filepath)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def _parse_csv(filepath: str) -> str:
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


def _parse_zip(filepath: str) -> Dict:
    text_parts = []
    sheets = []
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            with zipfile.ZipFile(filepath, "r") as zf:
                zf.extractall(tmpdir)
        except Exception as e:
            _get_logger().error(f"解压ZIP失败: {e}")
            return {"text_content": "", "sheets": []}

        for root, _dirs, files in os.walk(tmpdir):
            for fname in files:
                fpath = os.path.join(root, fname)
                ext = os.path.splitext(fname)[1].lower()
                if ext in (".xlsx", ".xls", ".pdf", ".docx", ".csv", ".txt"):
                    sub = parse_attachment(fpath, fname)
                    if sub["text_content"]:
                        text_parts.append(f"[{fname}] {sub['text_content']}")
                    if sub["sheets"]:
                        sheets.extend(sub["sheets"])
    return {"text_content": "\n".join(text_parts), "sheets": sheets}


def _parse_rar(filepath: str) -> Dict:
    text_parts = []
    sheets = []
    try:
        import rarfile
    except ImportError:
        _get_logger().warning("rarfile未安装，无法解析RAR附件")
        return {"text_content": "", "sheets": []}

    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            with rarfile.RarFile(filepath, "r") as rf:
                rf.extractall(tmpdir)
        except Exception as e:
            _get_logger().error(f"解压RAR失败: {e}")
            return {"text_content": "", "sheets": []}

        for root, _dirs, files in os.walk(tmpdir):
            for fname in files:
                fpath = os.path.join(root, fname)
                ext = os.path.splitext(fname)[1].lower()
                if ext in (".xlsx", ".xls", ".pdf", ".docx", ".csv", ".txt"):
                    sub = parse_attachment(fpath, fname)
                    if sub["text_content"]:
                        text_parts.append(f"[{fname}] {sub['text_content']}")
                    if sub["sheets"]:
                        sheets.extend(sub["sheets"])
    return {"text_content": "\n".join(text_parts), "sheets": sheets}


def extract_emails_from_text(text: str) -> List[str]:
    """从文本中提取邮箱地址"""
    if not text:
        return []
    pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    return list(set(re.findall(pattern, text)))


import re
