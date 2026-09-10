"""附件解析工具 — 支持 .zip .rar .xlsx .xls .pdf .docx .csv .jpg .png .jpeg .bmp"""
import os
import re
import tempfile
import zipfile
from typing import List, Dict

logger = None
_ocr_engine = None


def _get_logger():
    global logger
    if logger is None:
        import logging
        logger = logging.getLogger("mail_audit")
    return logger


_ocr_init_failed = False

# 图片附件持久化目录（懒加载 OCR 用，按内容 hash 去重）
ATTACHMENT_CACHE_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "cache", "attachments"
)


def save_image_attachment(payload: bytes, filename: str) -> str:
    """将图片附件持久化到 cache/attachments/, 返回落盘路径（按内容 hash 去重）"""
    import hashlib
    os.makedirs(ATTACHMENT_CACHE_DIR, exist_ok=True)
    h = hashlib.sha1(payload).hexdigest()[:12]
    safe = re.sub(r'[\\/:*?"<>|]', "_", filename) or "image"
    path = os.path.join(ATTACHMENT_CACHE_DIR, f"{h}_{safe}")
    if not os.path.exists(path):
        with open(path, "wb") as f:
            f.write(payload)
    return path


IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".tiff")

def _get_ocr():
    """懒加载 PaddleOCR 引擎（全程只初始化一次）"""
    global _ocr_engine, _ocr_init_failed
    if _ocr_init_failed:
        return None
    if _ocr_engine is None:
        try:
            # mkldnn 在部分 Windows 环境下触发段错误，必须在 paddle 导入前关闭
            os.environ.setdefault("FLAGS_use_mkldnn", "0")
            os.environ.setdefault("FLAGS_allocator_strategy", "auto_growth")
            from paddleocr import PaddleOCR
            try:
                # PaddleOCR 3.x：显式关闭 mkldnn
                _ocr_engine = PaddleOCR(lang="ch", enable_mkldnn=False)
            except TypeError:
                # PaddleOCR 2.x：不认识该参数，靠上面的环境变量兜底
                _ocr_engine = PaddleOCR(lang="ch")
            _get_logger().info("PaddleOCR 引擎初始化成功")
        except Exception as e:
            _get_logger().error(f"PaddleOCR 初始化失败: {e}")
            _ocr_init_failed = True
            return None
    return _ocr_engine


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
        if ext == ".xlsx":
            result.update(_parse_excel(filepath))
        elif ext == ".xls":
            result["text_content"] = _parse_xls(filepath)
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
        elif ext in IMAGE_EXTS:
            # 懒加载 OCR: 常规解析阶段不识别, 只登记待识别图片
            # OCR 兜底由 FieldExtractor._ocr_fallback 在字段仍缺失时触发
            result["text_content"] = ""
            result["ocr_pending"] = True
            result["filepath"] = filepath
        else:
            result["text_content"] = ""
            _get_logger().warning(f"不支持的附件格式: {filename}")
    except Exception as e:
        _get_logger().error(f"解析附件失败 {filename}: {e}")
        result["text_content"] = ""

    return result


def _parse_excel(filepath: str) -> Dict:
    import warnings
    from openpyxl import load_workbook
    text_parts = []
    sheets = []
    try:
        # openpyxl 在 load 和 iter_rows 期间都会对外部链接/扩展发 UserWarning;
        # 写到 stderr 时若控制台处于 QuickEdit 选择模式会永久阻塞 worker 线程
        # → 全程压制 (这些警告无业务价值)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
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


def _parse_xls(filepath: str) -> str:
    import xlrd
    text_parts = []
    wb = xlrd.open_workbook(filepath)
    for sheet in wb.sheets():
        for row_idx in range(min(sheet.nrows, 200)):
            cells = []
            for col_idx in range(sheet.ncols):
                val = sheet.cell_value(row_idx, col_idx)
                if val:
                    cells.append(str(val).strip())
            line = " ".join(cells)
            if line.strip():
                text_parts.append(line)
    return "\n".join(text_parts)


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


def _parse_image(filepath: str) -> str:
    """使用 PaddleOCR 识别图片中的文字（中英文）"""
    ocr = _get_ocr()
    if ocr is None:
        return ""
    try:
        result = ocr.ocr(filepath)
        return _extract_ocr_texts(result)
    except Exception as e:
        _get_logger().error(f"OCR识别失败 {filepath}: {e}")
        return ""


def _extract_ocr_texts(result) -> str:
    """兼容 PaddleOCR 2.x / 3.x 两种返回格式，抽成纯文本"""
    lines = []
    if not result:
        return ""
    for page in result:
        if not page:
            continue
        if isinstance(page, dict):
            # PaddleOCR 3.x: {'rec_texts': [...], 'rec_scores': [...]}
            for text in page.get("rec_texts") or []:
                if text and text.strip():
                    lines.append(text.strip())
        else:
            # PaddleOCR 2.x: [[box, (text, score)], ...]
            for line_info in page:
                try:
                    text = line_info[1][0]
                except (TypeError, IndexError):
                    continue
                if isinstance(text, str) and text.strip():
                    lines.append(text.strip())
    return "\n".join(lines)


def _parse_csv(filepath: str) -> str:
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


def _parse_zip(filepath: str) -> Dict:
    text_parts = []
    sheets = []
    pending_images = []
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            with zipfile.ZipFile(filepath, "r") as zf:
                zf.extractall(tmpdir)
        except Exception as e:
            _get_logger().error(f"解压ZIP失败: {e}")
            return {"text_content": "", "sheets": [], "pending_images": []}

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
                elif ext in IMAGE_EXTS:
                    # 压缩包内图片: 持久化后登记待 OCR（懒加载）
                    try:
                        with open(fpath, "rb") as f:
                            img_path = save_image_attachment(f.read(), fname)
                        pending_images.append({
                            "filename": fname, "text_content": "", "sheets": [],
                            "ocr_pending": True, "filepath": img_path,
                        })
                        text_parts.append(f"[{fname}]")
                    except Exception as e:
                        _get_logger().warning(f"压缩包内图片持久化失败 {fname}: {e}")
    return {"text_content": "\n".join(text_parts), "sheets": sheets,
            "pending_images": pending_images}


def _parse_rar(filepath: str) -> Dict:
    text_parts = []
    sheets = []
    pending_images = []
    with tempfile.TemporaryDirectory() as tmpdir:
        extracted = False

        # 方案1: 尝试 rarfile（需要 unrar.exe）
        try:
            import rarfile
            with rarfile.RarFile(filepath, "r") as rf:
                rf.extractall(tmpdir)
            extracted = True
        except Exception as e:
            _get_logger().warning(f"rarfile解压失败: {e}")

        # 方案2: 尝试 patool（纯Python调用命令行工具）
        if not extracted:
            try:
                import patool
                patool.extract_archive(filepath, outdir=tmpdir)
                extracted = True
                _get_logger().info(f"patool解压RAR成功: {filepath}")
            except Exception as e:
                _get_logger().warning(f"patool解压失败: {e}")

        # 方案3: 尝试 pyunpack + 7zip
        if not extracted:
            try:
                from pyunpack import Archive
                Archive(filepath).extractall(tmpdir)
                extracted = True
                _get_logger().info(f"pyunpack解压RAR成功: {filepath}")
            except Exception as e:
                _get_logger().warning(f"pyunpack解压失败: {e}")

        if not extracted:
            _get_logger().error(f"RAR解压全部失败，跳过: {filepath}")
            return {"text_content": "", "sheets": [], "pending_images": []}

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
                elif ext in IMAGE_EXTS:
                    # 压缩包内图片: 持久化后登记待 OCR（懒加载）
                    try:
                        with open(fpath, "rb") as f:
                            img_path = save_image_attachment(f.read(), fname)
                        pending_images.append({
                            "filename": fname, "text_content": "", "sheets": [],
                            "ocr_pending": True, "filepath": img_path,
                        })
                        text_parts.append(f"[{fname}]")
                    except Exception as e:
                        _get_logger().warning(f"压缩包内图片持久化失败 {fname}: {e}")
    return {"text_content": "\n".join(text_parts), "sheets": sheets,
            "pending_images": pending_images}


def extract_emails_from_text(text: str) -> List[str]:
    """从文本中提取邮箱地址"""
    if not text:
        return []
    pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    return list(set(re.findall(pattern, text)))


import re
