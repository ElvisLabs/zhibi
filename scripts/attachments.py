"""Plain text of a mail's attachments for the reply prompt (read-only; attachment content is untrusted data).

  extract(name, data) -> (kind, text)      # one file
  read_all(item) -> [{name, kind, size, text, truncated, note}]   # exchangelib item

Word / PowerPoint / plain text / mail are parsed with the stdlib; Excel needs openpyxl (xlrd for .xls),
PDF uses PyMuPDF, pypdf or pdftotext, whichever is there; .doc / .rtf go through macOS textutil.
Images are not read.
"""
import email
import html
import io
import re
import shutil
import subprocess
import tempfile
import zipfile
from email import policy
from pathlib import Path
from xml.etree import ElementTree as ET

MAX_BYTES = 20 * 1024 * 1024
PER_FILE = 20000
TOTAL = 40000
MAX_ROWS = 300

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
TEXT_EXT = {".txt", ".csv", ".tsv", ".md", ".log", ".json", ".xml", ".yaml", ".yml", ".ini", ".conf", ".cfg", ".py",
            ".java", ".kt", ".c", ".h", ".cpp", ".sh", ".diff", ".patch"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".heic", ".webp", ".svg", ".emf", ".wmf"}
KINDS = {".docx": "Word", ".docm": "Word", ".doc": "Word", ".rtf": "RTF", ".odt": "文档", ".pptx": "PPT",
         ".xlsx": "Excel", ".xlsm": "Excel", ".xls": "Excel", ".pdf": "PDF", ".eml": "邮件", ".zip": "压缩包",
         ".html": "网页", ".htm": "网页"}


class Unreadable(Exception):
    pass


def decode(data):
    for enc in ("utf-8-sig", "gb18030"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


def strip_html(text):
    text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", "", text)
    text = re.sub(r"(?i)<br\s*/?>|</(p|div|tr|li|h\d)>", "\n", text)
    text = re.sub(r"(?i)</t[dh]>", " | ", text)
    return html.unescape(re.sub(r"<[^>]+>", "", text))


def tidy(text):
    lines = [re.sub(r"[ \t 　]+", " ", l).strip() for l in text.replace("\r", "\n").split("\n")]
    out = []
    for l in lines:
        if l or (out and out[-1]):
            out.append(l)
    return "\n".join(out).strip()


# --- Word: body in order, tracked changes as [+inserted+] / [-deleted-], then comments ---

def docx_runs(el, state):
    parts = []
    for child in el:
        tag = child.tag
        if tag == W + "t":
            parts.append(child.text or "")
        elif tag == W + "delText":
            parts.append(child.text or "")
        elif tag in (W + "tab",):
            parts.append(" ")
        elif tag in (W + "br", W + "cr"):
            parts.append("\n")
        elif tag == W + "ins":
            inner = docx_runs(child, state)
            if inner.strip():
                state["ins"] += 1
                parts.append(f"[+{inner}+]")
        elif tag == W + "del":
            inner = docx_runs(child, state)
            if inner.strip():
                state["del"] += 1
                parts.append(f"[-{inner}-]")
        elif tag == W + "commentReference":
            parts.append(f"[批注{child.get(W + 'id')}]")
        elif tag not in (W + "pPr", W + "rPr", W + "instrText"):
            parts.append(docx_runs(child, state))
    return "".join(parts)


def docx_table(tbl, state):
    rows = []
    for tr in tbl.iter(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            text = " ".join(docx_runs(p, state).strip() for p in tc.iter(W + "p")).strip()
            if not cells or cells[-1] != text:  # merged cells repeat their text
                cells.append(text)
        rows.append(" | ".join(cells))
    return "\n".join(rows)


def docx_text(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    body = ET.fromstring(z.read("word/document.xml")).find(W + "body")
    state = {"ins": 0, "del": 0}
    out = []
    for el in body:
        if el.tag == W + "p":
            out.append(docx_runs(el, state))
        elif el.tag == W + "tbl":
            out.append(docx_table(el, state))
        elif el.tag == W + "sdt":
            out += [docx_runs(p, state) for p in el.iter(W + "p")]
    text = tidy("\n".join(out))
    if "word/comments.xml" in z.namelist():
        notes = []
        for c in ET.fromstring(z.read("word/comments.xml")).iter(W + "comment"):
            said = tidy("\n".join(docx_runs(p, state) for p in c.iter(W + "p")))
            notes.append(f"[批注{c.get(W + 'id')}] {c.get(W + 'author', '')}：{said}")
        if notes:
            text += "\n\n== 批注 ==\n" + "\n".join(notes)
    if state["ins"] or state["del"]:
        text = f"（文档含修订：插入 {state['ins']} 处 [+…+]，删除 {state['del']} 处 [-…-]）\n" + text
    return text


def pptx_text(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    slides = sorted((n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                    key=lambda n: int(re.search(r"\d+", n.rsplit("/", 1)[1]).group()))
    out = []
    for i, n in enumerate(slides, 1):
        paras = ["".join(t.text or "" for t in p.iter(A + "t")) for p in ET.fromstring(z.read(n)).iter(A + "p")]
        out.append(f"--- 第 {i} 页 ---\n" + "\n".join(p for p in paras if p.strip()))
    return "\n".join(out)


def xlsx_text(data):
    try:
        import openpyxl
    except ImportError:
        raise Unreadable("缺少 openpyxl")
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        rows = []
        for row in ws.iter_rows(values_only=True):
            cells = ["" if v is None else str(v).strip() for v in row]
            while cells and not cells[-1]:
                cells.pop()
            if cells:
                rows.append("\t".join(cells))
            if len(rows) >= MAX_ROWS:
                rows.append(f"…（只取前 {MAX_ROWS} 行）")
                break
        out.append(f"--- 工作表：{ws.title} ---\n" + "\n".join(rows))
    return "\n".join(out)


def xls_text(data):
    try:
        import xlrd
    except ImportError:
        raise Unreadable("缺少 xlrd")
    book = xlrd.open_workbook(file_contents=data)
    out = []
    for sh in book.sheets():
        rows = ["\t".join(str(v).strip() for v in sh.row_values(r)).rstrip() for r in range(min(sh.nrows, MAX_ROWS))]
        out.append(f"--- 工作表：{sh.name} ---\n" + "\n".join(r for r in rows if r))
    return "\n".join(out)


def pdf_text(data):
    try:
        import fitz

        with fitz.open(stream=data, filetype="pdf") as doc:
            return "\n".join(f"--- 第 {i} 页 ---\n{p.get_text()}" for i, p in enumerate(doc, 1))
    except ImportError:
        pass
    try:
        import pypdf

        reader = pypdf.PdfReader(io.BytesIO(data))
        return "\n".join(f"--- 第 {i} 页 ---\n{p.extract_text() or ''}" for i, p in enumerate(reader.pages, 1))
    except ImportError:
        pass
    if shutil.which("pdftotext"):
        return via_file(data, ".pdf", lambda f: ["pdftotext", "-layout", f, "-"])
    raise Unreadable("没有可用的 PDF 解析工具")


def via_file(data, suffix, cmd):
    with tempfile.TemporaryDirectory() as d:
        f = Path(d) / f"in{suffix}"
        f.write_bytes(data)
        r = subprocess.run(cmd(str(f)), capture_output=True, timeout=60)
        if r.returncode != 0:
            raise Unreadable(decode(r.stderr).strip()[:200] or "转换失败")
        return decode(r.stdout)


def eml_text(data):
    msg = email.message_from_bytes(data, policy=policy.default)
    head = "\n".join(f"{k}：{msg[k]}" for k in ("From", "To", "Cc", "Date", "Subject") if msg[k])
    part = msg.get_body(preferencelist=("plain", "html"))
    body = part.get_content() if part else ""
    if part and part.get_content_type() == "text/html":
        body = strip_html(body)
    names = [p.get_filename() for p in msg.iter_attachments() if p.get_filename()]
    return head + "\n\n" + body + (f"\n\n（附件：{'、'.join(names)}）" if names else "")


def extract(name, data):
    """(kind, text) for one attachment; raises Unreadable with a short reason."""
    ext = Path(name or "").suffix.lower()
    kind = KINDS.get(ext, ext.lstrip(".").upper() or "文件")
    if ext in IMAGE_EXT:
        raise Unreadable("图片，不读取内容")
    if len(data) > MAX_BYTES:
        raise Unreadable(f"文件过大（{len(data) / 1048576:.0f} MB）")
    try:
        if ext in (".docx", ".docm"):
            return kind, docx_text(data)
        if ext == ".pptx":
            return kind, pptx_text(data)
        if ext in (".xlsx", ".xlsm"):
            return kind, xlsx_text(data)
        if ext == ".xls":
            return kind, xls_text(data)
        if ext == ".pdf":
            return kind, pdf_text(data)
        if ext in (".doc", ".rtf", ".odt"):
            return kind, via_file(data, ext, lambda f: ["textutil", "-convert", "txt", "-stdout", f])
        if ext == ".eml":
            return kind, eml_text(data)
        if ext in (".html", ".htm"):
            return kind, strip_html(decode(data))
        if ext == ".zip":
            names = zipfile.ZipFile(io.BytesIO(data)).namelist()
            return kind, "压缩包内文件：\n" + "\n".join(names[:200])
        if ext in TEXT_EXT:
            return "文本", decode(data)
    except Unreadable:
        raise
    except Exception as exc:
        raise Unreadable(f"解析失败（{type(exc).__name__}）")
    raise Unreadable("格式不支持")


def read_all(item):
    """Every non-inline attachment of an exchangelib item, within PER_FILE / TOTAL characters."""
    out, budget = [], TOTAL
    for a in item.attachments or []:
        if getattr(a, "is_inline", False):
            continue
        entry = {"name": a.name or "", "kind": "", "size": getattr(a, "size", 0) or 0, "text": "", "truncated": False, "note": ""}
        out.append(entry)
        if budget <= 0:
            entry["note"] = "超出长度上限，未读取"
            continue
        try:
            if hasattr(a, "item"):  # an attached Outlook item
                it = a.item
                sender = getattr(it, "author", None) or getattr(it, "sender", None)
                entry["kind"] = "邮件"
                text = f"主题：{it.subject or ''}\n发件人：{getattr(sender, 'name', '') or ''}\n\n{it.text_body or ''}"
            else:
                if entry["size"] > MAX_BYTES:
                    raise Unreadable(f"文件过大（{entry['size'] / 1048576:.0f} MB）")
                entry["kind"], text = extract(a.name, a.content or b"")
        except Unreadable as exc:
            entry["kind"] = entry["kind"] or KINDS.get(Path(a.name or "").suffix.lower(), "文件")
            entry["note"] = str(exc)
            continue
        except Exception as exc:
            entry["note"] = f"读取失败（{type(exc).__name__}）"
            continue
        text = tidy(text)
        if not text:
            entry["note"] = "没有可读的文字（可能是扫描件）"
            continue
        limit = min(PER_FILE, budget)
        entry["text"], entry["truncated"] = text[:limit], len(text) > limit
        budget -= len(entry["text"])
    return out


if __name__ == "__main__":
    import json
    import sys

    for path in sys.argv[1:]:
        try:
            kind, text = extract(path, Path(path).read_bytes())
            print(json.dumps({"name": path, "kind": kind, "chars": len(tidy(text)), "head": tidy(text)[:1500]}, ensure_ascii=False, indent=1))
        except Unreadable as exc:
            print(json.dumps({"name": path, "note": str(exc)}, ensure_ascii=False))
