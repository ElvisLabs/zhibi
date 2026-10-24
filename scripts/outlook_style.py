"""Read the reply format Outlook for Mac would use: compose font, size, colour and signatures.

  python3 outlook_style.py [--refresh]     # prints JSON; cached in ~/Library/Caches/Zhibi/outlook_style.json

Sources, best first:
  - font / size / colour: the newest mail in Sent Items (what Outlook actually wrote; needs outlook-cli's
    python for exchangelib), else Outlook's defaults (Aptos 12pt) plus the colour in its preferences
  - signatures: the Outlook profile's Signatures folder; the default reply signature from its preferences
Stdlib only, except the Sent Items lookup.
"""
import argparse
import datetime
import html
import json
import plistlib
import re
import sqlite3
import sys
from pathlib import Path

CACHE = Path.home() / "Library/Caches/Zhibi/outlook_style.json"
PREFS = Path.home() / "Library/Containers/com.microsoft.Outlook/Data/Library/Preferences/com.microsoft.Outlook.plist"
PROFILE = Path.home() / "Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data"
OUTLOOK_DEFAULT = {"font_family": "Aptos, Arial, Helvetica, sans-serif", "font_size": "12pt"}
REGARDS_RE = re.compile(r"^\s*(best\s*regards|kind\s*regards|regards|br|thanks|cheers|此致|谢谢|祝好|顺祝)", re.I)


def prefs():
    try:
        return plistlib.loads(PREFS.read_bytes())
    except Exception:
        return {}


def hex_color(r, g, b):
    return "#%02X%02X%02X" % (round(r), round(g), round(b))


def default_color(p):
    raw = p.get("DefaultHTMLCompositionFontColor_R_G_B_Alpha")
    try:
        r, g, b = (float(x) for x in raw.split()[:3])
        return hex_color(r * 255, g * 255, b * 255)
    except Exception:
        return ""


def style_from_html(body):
    """First styled block of the sender's own text: Outlook stamps its compose font on every line."""
    m = re.search(r"<body[^>]*>(.*)", body or "", re.S | re.I)
    for style in re.findall(r'<(?:div|p|span)[^>]*style="([^"]*font-family[^"]*)"', m.group(1) if m else ""):
        decl = dict((k.strip().lower(), v.strip()) for k, _, v in (d.partition(":") for d in style.split(";")) if v)
        if "font-family" not in decl:
            continue
        color = decl.get("color", "")
        rgb = re.match(r"rgb\((\d+),\s*(\d+),\s*(\d+)\)", color)
        return {
            "font_family": html.unescape(decl["font-family"]),
            "font_size": decl.get("font-size", ""),
            "color": hex_color(*map(int, rgb.groups())) if rgb else color,
        }
    return None


def sent_mail_style():
    try:
        from outlook_cli.exchange import get_account
    except ImportError:
        return None
    try:
        account = get_account()
        for item in account.sent.all().order_by("-datetime_sent").only("body", "subject", "datetime_sent")[:5]:
            style = style_from_html(str(item.body or ""))
            if style:
                style["from_mail"] = f"{item.datetime_sent:%Y-%m-%d} {item.subject or ''}"[:80]
                return style
    except Exception:
        return None
    return None


def clean_signature_html(raw):
    """Word-generated signature → inline-styled fragment that survives a paste."""
    body = re.search(r"<body[^>]*>(.*)</body>", raw, re.S | re.I)
    s = body.group(1) if body else raw
    s = re.sub(r"<o:p>.*?</o:p>", "", s, flags=re.S)
    if re.search(r"<div[^>]*class=\"?WordSection1", s):
        s = re.sub(r"<div[^>]*class=\"?WordSection1\"?[^>]*>", "", s, count=1)
        s = re.sub(r"</div>\s*$", "", s.strip())
    s = re.sub(r"<p class=\"?MsoNormal\"?(\s+style='([^']*)')?", lambda m: f"<p style='margin:0;{m.group(2) or ''}'", s)
    s = re.sub(r"\sclass=\"?Mso\w+\"?", "", s)
    # Word ends signatures with an empty paragraph.
    empty = r"(<div>)?\s*<p[^>]*>(\s|&nbsp;|<span[^>]*>(\s|&nbsp;)*</span>)*</p>\s*(</div>)?\s*$"
    while re.search(empty, s.strip()):
        s = re.sub(empty, "", s.strip())
    return s.strip()


def signature_images(con, record_id):
    """{content-id: data URI}. Images sit in 'Signature Attachments' as MIME parts, linked by GUID
    through Signatures_OwnedBlocks."""
    images = {}
    try:
        rows = con.execute("select BlockID from Signatures_OwnedBlocks where Record_RecordID=?", (record_id,)).fetchall()
    except Exception:
        return images
    for (block,) in rows:
        if not isinstance(block, bytes) or len(block) < 20:
            continue
        h = block[4:20].hex().upper()
        guid = f"{h[:8]}-{h[8:12]}-{h[12:16]}-{h[16:20]}-{h[20:]}"
        for f in (PROFILE / "Signature Attachments").glob(f"*/{guid}.olk15SigAttachment"):
            text = f.read_bytes().decode("latin-1")
            mime = re.search(r"Content-type:\s*([\w/+.-]+)", text, re.I)
            cid = re.search(r"Content-ID:\s*<([^>]+)>", text, re.I)
            body = re.search(r"Content-transfer-encoding:\s*base64\s*?(?:\r\n|\r|\n){2}([A-Za-z0-9+/=\r\n]+)", text, re.I)
            if mime and cid and body:
                data = re.sub(r"\s+", "", body.group(1))
                images[cid.group(1)] = f"data:{mime.group(1)};base64,{data}"
    return images


def signatures():
    out = []
    db = PROFILE / "Outlook.sqlite"
    try:
        con = sqlite3.connect(f"file:{db}?immutable=1", uri=True)
        rows = con.execute("select Record_RecordID, PathToDataFile from Signatures").fetchall()
    except Exception:
        return out
    for record_id, rel in rows:
        f = PROFILE / rel
        try:
            raw = f.read_bytes()
        except OSError:
            continue
        start = raw.find("<html".encode("utf-16le"))
        end = raw.find("</html>".encode("utf-16le"), start)
        if start < 0 or end < 0:
            continue
        doc = raw[start:end + 14].decode("utf-16le", "replace")
        # The name is the UTF-16 text right before the HTML, after any image file names.
        lead = raw[max(0, start - 200):start]
        lead = lead[len(lead) % 2:].decode("utf-16le", "replace")
        name = re.split(r"[\x00-\x1f]|\.(?:jpg|jpeg|png|gif)", lead, flags=re.I)[-1].strip() or f.stem[:8]
        frag = clean_signature_html(doc)
        imgs = signature_images(con, record_id)
        cids = re.findall(r'src="cid:([^"]+)"', frag)
        for cid in cids:
            if cid in imgs:
                frag = frag.replace(f'src="cid:{cid}"', f'src="{imgs[cid]}"')
        text = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", frag))).strip()
        out.append({
            "id": f.stem, "name": name, "html": frag, "text": text,
            "has_images": bool(cids), "images_found": all(c in imgs for c in cids),
            "starts_with_regards": bool(REGARDS_RE.match(text)),
        })
    return out


def default_signatures(p, sigs):
    """Best effort: Outlook keeps per-account defaults under DefaultSignatures (absent when none is set)."""
    raw = json.dumps(p.get("DefaultSignatures", {}), default=str)
    found = {"new": None, "reply": None}
    for kind, pattern in (("reply", r"(?i)(reply|fwd|forward)[^\"]*\"\s*:\s*\"([^\"]+)"), ("new", r"(?i)new[^\"]*\"\s*:\s*\"([^\"]+)")):
        for m in re.finditer(pattern, raw):
            value = m.group(m.lastindex)
            hit = next((s["id"] for s in sigs if s["id"].lower() in value.lower() or s["name"] == value), None)
            if hit:
                found[kind] = hit
                break
    return found


def detect(refresh=False, max_age_hours=24):
    if not refresh and CACHE.exists():
        try:
            data = json.loads(CACHE.read_text(encoding="utf-8"))
            age = datetime.datetime.now() - datetime.datetime.fromisoformat(data["detected_at"])
            if age < datetime.timedelta(hours=max_age_hours):
                return data
        except Exception:
            pass
    p = prefs()
    style = sent_mail_style()
    if style:
        style["source"] = "最近一封已发邮件"
    else:
        style = dict(OUTLOOK_DEFAULT, color=default_color(p), source="Outlook 默认设置")
    sigs = signatures()
    data = {
        "detected_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "font_family": style.get("font_family") or OUTLOOK_DEFAULT["font_family"],
        "font_size": style.get("font_size") or OUTLOOK_DEFAULT["font_size"],
        "color": style.get("color", ""),
        "source": style["source"],
        "from_mail": style.get("from_mail", ""),
        "signatures": sigs,
        "default_signature": default_signatures(p, sigs),
    }
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return data


def cached():
    """Cache only (for render.py, which must stay fast and stdlib-only)."""
    try:
        return json.loads(CACHE.read_text(encoding="utf-8"))
    except Exception:
        return None


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--refresh", action="store_true", help="Ignore the cache and read Outlook again")
    args = ap.parse_args()
    d = detect(args.refresh)
    for s in d["signatures"]:
        s["html"] = s["html"][:300] + ("…" if len(s["html"]) > 300 else "")
    print(json.dumps(d, ensure_ascii=False, indent=1))
