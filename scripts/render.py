#!/usr/bin/env python3
"""Render a Markdown-subset mail body into Outlook-safe, inline-styled HTML.

Usage:
  render.py BODY.md [-o OUT_BASE] [--meta info.json] [--no-signature] [--open]

Writes OUT_BASE.html (fragment to send) and OUT_BASE.preview.html (full page
with recipients header and the original mail), then prints a JSON summary.
"""
import argparse
import datetime
import html
import json
import re
import subprocess
import sys
from pathlib import Path

# Shared by the 执笔 app and the skill; settings saved from the app's pane land here too.
SUPPORT_DIR = Path.home() / "Library/Application Support/Zhibi"
CONFIG_PATH = SUPPORT_DIR / "config.json"
STYLE_CACHE = Path.home() / "Library/Caches/Zhibi/outlook_style.json"

DEFAULTS = {
    "closing": "BR",
    "signature_lines": [],
    "font_family": "'Microsoft YaHei','微软雅黑','PingFang SC','Helvetica Neue',Arial,sans-serif",
    "font_size": "10.5pt",
    "text_color": "#1f2329",
    "accent_color": "#2F5496",
    "muted_color": "#646a73",
    "border_color": "#d0d7de",
    "header_bg": "#eef2f8",
}

STATUS_COLORS = [
    (r"(已完成|完成|已解决|已关闭|已合入|已上线|通过|正常|done|ok|pass(ed)?|closed|fixed|resolved)", "#1a7f37"),
    (r"(进行中|处理中|开发中|测试中|验证中|分析中|in ?progress|wip|ongoing)", "#0969da"),
    (r"(风险|高风险|延期|延误|阻塞|失败|未通过|超期|异常|ng|fail(ed)?|blocked|delayed|at risk)", "#cf222e"),
    (r"(待定|待确认|未开始|待排期|暂停|挂起|tbd|pending|on hold|not started)", "#9a6700"),
]
STATUS_RES = [(re.compile(r"^[\s🟢🟡🔴⚪🔵✅❌⚠️]*" + p + r"\s*$", re.I), c) for p, c in STATUS_COLORS]

PLACEHOLDER_RE = re.compile(r"【待确认[:：]?[^】]*】")
LIST_RE = re.compile(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$")
TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")
HR_RE = re.compile(r"^\s*(-{3,}|\*{3,}|_{3,})\s*$")
HEADING_RE = re.compile(r"^(#{1,3})\s+(.*?)\s*#*\s*$")


def load_config():
    cfg = dict(DEFAULTS)
    if CONFIG_PATH.exists():
        cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    apply_outlook_format(cfg)
    return cfg


def apply_outlook_format(cfg):
    """Font, size, colour and signature from Outlook (cached by outlook_style.py) unless config says custom."""
    try:
        o = json.loads(STYLE_CACHE.read_text(encoding="utf-8"))
    except Exception:
        o = {}
    if cfg.get("format_source", "outlook") == "outlook" and o.get("font_family"):
        cfg["font_family"] = o["font_family"]
        cfg["font_size"] = o.get("font_size") or cfg["font_size"]
        cfg["text_color"] = o.get("color") or cfg["text_color"]
    sigs = o.get("signatures") or []
    choice = cfg.get("signature", "outlook")
    cfg["_signature_html"], cfg["_skip_closing"] = "", False
    if choice == "outlook":
        # Outlook inserts its default reply signature into the reply window itself.
        sid = (o.get("default_signature") or {}).get("reply")
        sig = next((x for x in sigs if x["id"] == sid), None)
        cfg["_skip_closing"] = bool(sig and sig.get("starts_with_regards"))
    elif choice and choice != "none":
        sig = next((x for x in sigs if choice in (x["name"], x["id"])), None)
        if sig:
            cfg["_signature_html"] = sig["html"]
            cfg["_skip_closing"] = bool(sig.get("starts_with_regards"))


def points(size):
    m = re.match(r"\s*([\d.]+)\s*(pt|px)?", str(size))
    if not m:
        return 10.5
    v = float(m.group(1))
    return v * 0.75 if m.group(2) == "px" else v


class Renderer:
    def __init__(self, cfg):
        self.cfg = cfg
        self.placeholders = []
        self.base = points(cfg["font_size"])

    def pt(self, delta):
        """A size relative to the body font, so headings and tables scale with Outlook's font size."""
        return f"{round(self.base + delta, 1):g}pt"

    # --- inline ---

    def inline(self, text):
        stash = []

        def keep(fragment):
            stash.append(fragment)
            return f"\x00{len(stash) - 1}\x00"

        mono = f"font-family:Consolas,Menlo,'Courier New',monospace;font-size:{self.pt(-1)};background:#f3f4f6;padding:0 3px;"
        text = re.sub(r"`([^`]+)`", lambda m: keep(f'<span style="{mono}">{html.escape(m.group(1))}</span>'), text)

        def link(m):
            label, url = m.group(1), m.group(2)
            return keep(f'<a href="{html.escape(url, quote=True)}" style="color:#0969da;text-decoration:underline;">{html.escape(label)}</a>')

        text = re.sub(r"\[([^\]]+)\]\((\S+?)\)", link, text)
        text = re.sub(
            r"(?<![\"'=])(https?://[^\s<>（）()，。；]+)",
            lambda m: keep(f'<a href="{html.escape(m.group(1), quote=True)}" style="color:#0969da;">{html.escape(m.group(1))}</a>'),
            text,
        )

        def placeholder(m):
            self.placeholders.append(m.group(0))
            return keep(f'<span style="background:#ffebe9;color:#cf222e;font-weight:bold;">{html.escape(m.group(0))}</span>')

        text = PLACEHOLDER_RE.sub(placeholder, text)
        text = html.escape(text, quote=False)
        text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
        text = re.sub(r"~~(.+?)~~", r"<s>\1</s>", text)
        text = re.sub(r"==(.+?)==", r'<span style="background:#fff3b0;">\1</span>', text)
        return re.sub(r"\x00(\d+)\x00", lambda m: stash[int(m.group(1))], text)

    def lines_inline(self, lines):
        return "<br>".join(self.inline(l.strip()) for l in lines)

    # --- blocks ---

    def p(self, content, extra="", margin="0 0 8px 0"):
        return f'<p style="margin:{margin};line-height:1.6;{extra}">{content}</p>'

    def heading(self, level, text):
        c = self.cfg
        if level == 1:
            return self.p(f"<b>{self.inline(text)}</b>", f"font-size:{self.pt(1.5)};color:{c['accent_color']};", "14px 0 8px 0")
        if level == 2:
            return self.p(
                f"<b>{self.inline(text)}</b>",
                f"font-size:{self.pt(0.5)};border-left:3px solid {c['accent_color']};padding-left:8px;",
                "14px 0 6px 0",
            )
        return self.p(f"<b>{self.inline(text)}</b>", "", "10px 0 4px 0")

    def callout(self, lines):
        c = self.cfg
        inner = self.render("\n".join(lines))
        return (
            '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:2px 0 10px 0;"><tr>'
            f'<td style="background:#f4f7fc;border-left:3px solid {c["accent_color"]};padding:8px 12px 0 12px;">{inner}</td>'
            "</tr></table>"
        )

    def code_block(self, lines):
        body = "<br>".join(html.escape(l).replace(" ", "&nbsp;") or "&nbsp;" for l in lines)
        return (
            '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:2px 0 10px 0;"><tr>'
            '<td style="background:#f6f8fa;border:1px solid #d0d7de;padding:8px 12px;'
            f"font-family:Consolas,Menlo,'Courier New',monospace;font-size:{self.pt(-1.5)};line-height:1.5;\">{body}</td>"
            "</tr></table>"
        )

    @staticmethod
    def split_row(line):
        line = line.strip()
        if line.startswith("|"):
            line = line[1:]
        if line.endswith("|") and not line.endswith("\\|"):
            line = line[:-1]
        cells = re.split(r"(?<!\\)\|", line)
        return [cell.strip().replace("\\|", "|") for cell in cells]

    def table(self, header, sep, rows):
        c = self.cfg
        aligns = []
        for spec in self.split_row(sep):
            if spec.startswith(":") and spec.endswith(":"):
                aligns.append("center")
            elif spec.endswith(":"):
                aligns.append("right")
            else:
                aligns.append("left")
        width = len(header)
        aligns = (aligns + ["left"] * width)[:width]
        cell = f"border:1px solid {c['border_color']};padding:5px 10px;vertical-align:top;line-height:1.5;"
        out = [f'<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:2px 0 10px 0;font-size:{self.pt(-0.5)};">']
        out.append("<tr>")
        for i, h in enumerate(header):
            out.append(
                f'<th style="{cell}background:{c["header_bg"]};font-weight:bold;text-align:{aligns[i]};white-space:nowrap;">{self.inline(h)}</th>'
            )
        out.append("</tr>")
        for row in rows:
            row = (row + [""] * width)[:width]
            out.append("<tr>")
            for i, value in enumerate(row):
                style = f"{cell}text-align:{aligns[i]};"
                for regex, color in STATUS_RES:
                    if value and regex.match(value):
                        style += f"color:{color};font-weight:bold;white-space:nowrap;"
                        break
                out.append(f'<td style="{style}">{self.inline(value).replace("&lt;br&gt;", "<br>")}</td>')
            out.append("</tr>")
        out.append("</table>")
        return "".join(out)

    def list_block(self, items):
        out, stack = [], []

        def open_list(tag, num, depth):
            start = f' start="{num}"' if tag == "ol" and num and num != 1 else ""
            margin = "0 0 8px 0" if depth == 0 else "2px 0 2px 0"
            kind = "" if tag == "ol" else ("list-style-type:disc;" if depth == 0 else "list-style-type:circle;")
            out.append(f'<{tag}{start} style="margin:{margin};padding-left:22px;{kind}">')

        for indent, tag, num, text in items:
            while stack and indent < stack[-1][0]:
                out.append(f"</li></{stack.pop()[1]}>")
            if stack and indent == stack[-1][0]:
                if stack[-1][1] != tag:
                    out.append(f"</li></{stack.pop()[1]}>")
                    open_list(tag, num, len(stack))
                    stack.append((indent, tag))
                else:
                    out.append("</li>")
            else:
                open_list(tag, num, len(stack))
                stack.append((indent, tag))
            out.append(f'<li style="margin:0 0 3px 0;line-height:1.6;">{self.lines_inline(text.split(chr(10)))}')
        while stack:
            out.append(f"</li></{stack.pop()[1]}>")
        return "".join(out)

    def render(self, text):
        lines = text.replace("\r\n", "\n").replace("\t", "    ").split("\n")
        blocks, para, i, n = [], [], 0, len(lines)

        def flush():
            if para:
                blocks.append(self.p(self.lines_inline(para)))
                para.clear()

        while i < n:
            line = lines[i]
            stripped = line.strip()
            if not stripped:
                flush()
                i += 1
                continue
            if stripped.startswith("```"):
                flush()
                i += 1
                code = []
                while i < n and not lines[i].strip().startswith("```"):
                    code.append(lines[i])
                    i += 1
                blocks.append(self.code_block(code))
                i += 1
                continue
            m = HEADING_RE.match(stripped)
            if m:
                flush()
                blocks.append(self.heading(len(m.group(1)), m.group(2)))
                i += 1
                continue
            if HR_RE.match(stripped):
                flush()
                blocks.append(f'<hr style="border:none;border-top:1px solid {self.cfg["border_color"]};margin:12px 0;">')
                i += 1
                continue
            if stripped.startswith(">"):
                flush()
                quote = []
                while i < n and lines[i].strip().startswith(">"):
                    quote.append(re.sub(r"^\s*>\s?", "", lines[i]))
                    i += 1
                blocks.append(self.callout(quote))
                continue
            if stripped.startswith("|") and i + 1 < n and TABLE_SEP_RE.match(lines[i + 1]):
                flush()
                header, sep = self.split_row(line), lines[i + 1]
                i += 2
                rows = []
                while i < n and lines[i].strip().startswith("|"):
                    rows.append(self.split_row(lines[i]))
                    i += 1
                blocks.append(self.table(header, sep, rows))
                continue
            if LIST_RE.match(line):
                flush()
                items = []
                while i < n:
                    lm = LIST_RE.match(lines[i])
                    if lm:
                        marker = lm.group(2)
                        ordered = marker[0].isdigit()
                        items.append([len(lm.group(1)), "ol" if ordered else "ul", int(marker[:-1]) if ordered else None, lm.group(3)])
                    elif lines[i].strip() and lines[i].startswith("  ") and items:
                        items[-1][3] += "\n" + lines[i].strip()
                    else:
                        break
                    i += 1
                blocks.append(self.list_block(items))
                continue
            para.append(line)
            i += 1
        flush()
        return "".join(blocks)

    def signature(self):
        c = self.cfg
        out = []
        if c.get("closing") and not c.get("_skip_closing"):
            out.append(self.p(self.inline(c["closing"]), "", "12px 0 8px 0"))
        if c.get("_signature_html"):
            # The user's own Outlook signature, already inline-styled.
            out.append(f'<div style="margin-top:{"12px" if c.get("_skip_closing") else "4px"};line-height:normal;">{c["_signature_html"]}</div>')
        lines = c.get("signature_lines") or []
        if lines:
            sig = "<br>".join(self.inline(l) for l in lines)
            out.append(self.p(sig, f"font-size:{self.pt(-1.5)};color:{c['muted_color']};"))
        return "".join(out)

    def wrap(self, inner):
        c = self.cfg
        return (
            f'<div style="font-family:{c["font_family"]};font-size:{c["font_size"]};color:{c["text_color"]};line-height:1.6;">'
            f"{inner}</div><p style=\"margin:0;\">&nbsp;</p>"
        )


# Preview only: Macs often ship a regular-only YaHei, so map it to PingFang to show bold as Windows will.
PREVIEW_FONT_ALIAS = "".join(
    f"@font-face{{font-family:'{family}';src:local('{face}');font-weight:{weight};}}"
    for family in ("Microsoft YaHei", "微软雅黑")
    for face, weight in (("PingFangSC-Regular", 400), ("PingFangSC-Semibold", 700))
)


def fmt_people(people):
    out = []
    for p in people or []:
        if isinstance(p, dict):
            name, email = p.get("name") or "", p.get("email") or ""
            out.append(f"{name} &lt;{html.escape(email)}&gt;" if name and name != email else html.escape(email))
        else:
            out.append(html.escape(str(p)))
    return "; ".join(out) or "—"


def preview_page(fragment, meta, cfg):
    rows = []
    if meta:
        rows = [
            ("模式", html.escape(meta.get("mode", ""))),
            ("主题", html.escape(meta.get("reply_subject") or meta.get("subject") or "")),
            ("收件人", fmt_people(meta.get("to"))),
            ("抄送", fmt_people(meta.get("cc"))),
        ]
        if meta.get("attachments"):
            rows.append(("附件", html.escape(", ".join(meta["attachments"]))))
    head = "".join(
        f'<tr><td style="color:#646a73;padding:2px 12px 2px 0;white-space:nowrap;vertical-align:top;">{k}</td><td>{v}</td></tr>'
        for k, v in rows
    )
    original = ""
    if meta and meta.get("original"):
        o = meta["original"]
        body = html.escape((o.get("body") or "").strip())
        original = (
            '<div style="margin-top:16px;border-top:1px solid #d0d7de;padding-top:10px;color:#646a73;font-size:9.5pt;">'
            f"<b>原邮件</b> · {html.escape(o.get('sender_name') or '')} &lt;{html.escape(o.get('sender') or '')}&gt; · {html.escape(o.get('date') or '')}<br>"
            f"<b>主题：</b>{html.escape(o.get('subject') or '')}"
            f'<pre style="white-space:pre-wrap;font-family:inherit;margin:8px 0 0 0;">{body}</pre></div>'
        )
    return (
        '<!doctype html><html><head><meta charset="utf-8"><title>邮件预览</title>'
        f"<style>{PREVIEW_FONT_ALIAS}</style></head>"
        '<body style="background:#f5f6f7;margin:0;padding:24px;">'
        '<div style="max-width:860px;margin:0 auto;background:#fff;border:1px solid #e5e6eb;border-radius:6px;padding:20px 24px;">'
        f'<table style="font-family:{cfg["font_family"]};font-size:10pt;margin-bottom:12px;border-bottom:1px solid #e5e6eb;padding-bottom:8px;width:100%;">{head}</table>'
        f"{fragment}{original}</div></body></html>"
    )


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", help="Markdown body file ('-' for stdin)")
    ap.add_argument("-o", "--out", help="Output base path (default: source path without .md)")
    ap.add_argument("--meta", help="JSON from mail_info.py info, shown in the preview header")
    ap.add_argument("--no-signature", action="store_true", help="Skip closing/signature (digests, forwards)")
    ap.add_argument("--open", action="store_true", help="Open the preview in the default browser")
    args = ap.parse_args()

    text = sys.stdin.read() if args.source == "-" else Path(args.source).read_text(encoding="utf-8")
    if args.out:
        base = Path(args.out)
    elif args.source != "-":
        base = Path(args.source).with_suffix("")
    else:
        base = Path("_ai/reports/mail") / datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    base.parent.mkdir(parents=True, exist_ok=True)

    cfg = load_config()
    r = Renderer(cfg)
    inner = r.render(text)
    if not args.no_signature:
        inner += r.signature()
    fragment = r.wrap(inner)

    meta = None
    if args.meta:
        meta = json.loads(Path(args.meta).read_text(encoding="utf-8"))
        meta = meta.get("data", meta)

    html_path = base.with_name(base.name + ".html")
    preview_path = base.with_name(base.name + ".preview.html")
    html_path.write_text(fragment, encoding="utf-8")
    preview_path.write_text(preview_page(fragment, meta, cfg), encoding="utf-8")
    if args.open:
        subprocess.run(["open", str(preview_path)], check=False)

    print(json.dumps(
        {
            "ok": True,
            "html": str(html_path.resolve()),
            "preview": str(preview_path.resolve()),
            "placeholders": r.placeholders,
            "chars": len(fragment),
        },
        ensure_ascii=False,
    ))


if __name__ == "__main__":
    main()
