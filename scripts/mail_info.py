"""Read one mail through outlook-cli's login and print what a reply needs (read-only).

Run with outlook-cli's own interpreter so its exchangelib + keyring login are reused:
  PY=$(head -1 "$(command -v outlook-cli)" | sed 's/^#!//')
  $PY mail_info.py info --id <mail_id> [--all] [--body-chars 6000] [--attachments]
  $PY mail_info.py info --subject "<exact subject>" --sender <email> [--folder 收件箱] [--all]

Adds what `outlook-cli mail read` lacks: display names, reply / reply-all recipients
without yourself, the reply subject, times in outlook-cli's configured timezone and, with
--attachments, the text of the attachments (attachments.py).
The --subject form finds the newest matching mail (what Outlook shows as selected).
"""
import argparse
import json
import re
import sys

import attachments

try:
    from outlook_cli import __version__ as OC_VERSION
    from outlook_cli.config import get_permission_mode
    from outlook_cli.config import load as load_oc_config
    from outlook_cli.exchange import find_mail_by_id, get_account
except ImportError as exc:
    print(json.dumps({"ok": False, "error": {"code": "E_CONFIG", "message": f"run with a python that has outlook-cli: {exc}"}}))
    sys.exit(4)

REPLY_PREFIX_RE = re.compile(r"^\s*(re|回复|答复)\s*[:：]", re.I)
ANY_PREFIX_RE = re.compile(r"^\s*((re|fw|fwd|回复|答复|转发)\s*[:：]\s*)+", re.I)


class MailError(Exception):
    def __init__(self, message, code="E_VALIDATION", exit_code=2):
        super().__init__(message)
        self.code, self.exit_code = code, exit_code


def emit(data, ok=True, code=0):
    print(json.dumps({"ok": ok, "data" if ok else "error": data}, ensure_ascii=False, indent=None))
    sys.exit(code)


def fail(message, code="E_VALIDATION", exit_code=2, **details):
    emit({"code": code, "message": message, **details}, ok=False, code=exit_code)


def norm(text):
    return " ".join((text or "").replace(" ", " ").split())


def core(subject):
    return ANY_PREFIX_RE.sub("", norm(subject))


def local_time(dt):
    from zoneinfo import ZoneInfo

    # Show times in outlook-cli's configured zone, not the laptop's current one.
    tz = load_oc_config().get("timezone")
    return dt.astimezone(ZoneInfo(tz) if tz else None).strftime("%Y-%m-%d %H:%M")


def person(mb):
    return {"name": getattr(mb, "name", "") or "", "email": (getattr(mb, "email_address", "") or "").lower()}


def reply_subject(subject):
    subject = subject or ""
    return subject if REPLY_PREFIX_RE.match(subject) else f"RE: {subject}"


def recipients(item, me, reply_all):
    sender = item.author or item.sender
    to = [person(sender)] if sender else []
    cc = []
    if reply_all:
        seen = {p["email"] for p in to} | {me}
        for mb in item.to_recipients or []:
            p = person(mb)
            if p["email"] not in seen:
                to.append(p)
                seen.add(p["email"])
        for mb in item.cc_recipients or []:
            p = person(mb)
            if p["email"] not in seen:
                cc.append(p)
                seen.add(p["email"])
    return to, cc


def folder_name(account, item):
    pid = getattr(item.parent_folder_id, "id", None)
    if not pid:
        return ""
    try:
        return account.root.get_folder(item.parent_folder_id).name
    except Exception:
        return next((f.name for f in account.msg_folder_root.walk() if f.id == pid), "")


def candidate_folders(account, name):
    standard = [account.inbox, account.sent]
    name = norm(name)
    if not name:
        return standard
    picked = next((f for f in standard + [account.drafts, account.trash, account.junk] if f.name == name), None)
    if picked is None:
        try:
            picked = next((f for f in account.msg_folder_root.walk() if f.name == name), None)
        except Exception:
            picked = None
    return [picked] + [f for f in standard if f is not picked] if picked else standard


def sender_email(item):
    mb = item.author or item.sender
    return (getattr(mb, "email_address", "") or "").lower()


def find_by_subject(account, subject, sender, folder):
    """Newest mail whose subject equals `subject` (and sender matches), searching `folder` first."""
    want, sender = norm(subject), (sender or "").lower()
    for f in candidate_folders(account, folder):
        for query in ({"subject": subject}, {"subject__contains": core(subject) or want}):
            try:
                hits = f.filter(**query).order_by("-datetime_received").only("subject", "sender", "author")[:30]
                for it in hits:
                    if norm(it.subject) == want and (not sender or sender_email(it) == sender):
                        return f.get(id=it.id)
            except Exception:
                continue
    raise MailError(f"Mail not found: {subject!r} from {sender or '?'}", "E_NOT_FOUND", 3)


def load_item(args):
    account = get_account()
    if args.id:
        item = find_mail_by_id(account, args.id)
        if not item:
            raise MailError(f"Mail not found: {args.id}", "E_NOT_FOUND", 3)
        return account, item
    if not args.subject:
        raise MailError("give --id or --subject")
    return account, find_by_subject(account, args.subject, args.sender, args.folder)


def build_info(account, item, reply_all, body_chars=6000, with_attachments=False):
    me = account.primary_smtp_address.lower()
    to, cc = recipients(item, me, reply_all)
    body = item.text_body or ""
    sender = item.author or item.sender
    everyone = [person(m) for m in (item.to_recipients or []) + (item.cc_recipients or [])]
    return {
        "mode": "reply-all" if reply_all else "reply",
        "id": item.id,
        "subject": item.subject or "",
        "reply_subject": reply_subject(item.subject),
        "to": to,
        "cc": cc,
        "me": me,
        "me_name": next((p["name"] for p in everyone if p["email"] == me and p["name"]), ""),
        "permission_mode": get_permission_mode(),
        "outlook_cli_version": OC_VERSION,
        "original": {
            "subject": item.subject or "",
            "sender": person(sender)["email"] if sender else "",
            "sender_name": person(sender)["name"] if sender else "",
            "to": [person(m) for m in item.to_recipients or []],
            "cc": [person(m) for m in item.cc_recipients or []],
            "date": local_time(item.datetime_received) if item.datetime_received else "",
            "folder": folder_name(account, item),
            "attachments": [a.name for a in item.attachments or [] if not getattr(a, "is_inline", False)],
            "body": body[:body_chars],
            "body_truncated": len(body) > body_chars,
            "attachment_files": attachments.read_all(item) if with_attachments else [],
        },
        "_untrusted": ["subject", "to", "cc", "original"],
    }


def cmd_info(args):
    account, item = load_item(args)
    emit(build_info(account, item, args.all, args.body_chars, args.attachments))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("info", help="Original mail + computed reply recipients (read-only)")
    p.add_argument("--id")
    p.add_argument("--subject", help="Exact subject as shown in Outlook (instead of --id)")
    p.add_argument("--sender", help="Sender email, narrows --subject lookup")
    p.add_argument("--folder", help="Folder display name to search first, e.g. 收件箱")
    p.add_argument("--all", action="store_true", help="Reply-all recipients")
    p.add_argument("--body-chars", type=int, default=6000)
    p.add_argument("--attachments", action="store_true", help="Also read the attachments' text")
    p.set_defaults(fn=cmd_info)

    args = ap.parse_args()
    try:
        args.fn(args)
    except SystemExit:
        raise
    except MailError as exc:
        fail(str(exc), exc.code, exc.exit_code)
    except Exception as exc:
        fail(f"{type(exc).__name__}: {exc}", "E_SERVER", 7)


if __name__ == "__main__":
    main()
