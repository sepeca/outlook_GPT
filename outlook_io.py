
import datetime as dt
from typing import List, Optional
from models import EmailItem
from utils import html_to_text, extract_signature, to_naive_local
def _resolve_folder(ns, folder_path: str):
    """Resolve an Outlook folder by path. First segment may be a store name or a well-known folder."""
    if not folder_path:
        return ns.GetDefaultFolder(6)  # Inbox
    parts = [p for p in str(folder_path).split("/") if p]
    if not parts:
        return ns.GetDefaultFolder(6)

    first = parts[0].lower()
    # Common localized Inbox names
    inbox_aliases = {"inbox", "входящие", "doručená pošta", "prijata posta", "posta doručena"}
    if first in inbox_aliases:
        folder = ns.GetDefaultFolder(6)  # olFolderInbox
        start_idx = 1
    else:
        # Try to match a store/mailbox by name
        folder = None
        for i in range(1, ns.Folders.Count + 1):
            root = ns.Folders.Item(i)
            if str(root.Name).lower() == first:
                folder = root
                break
        if folder is None:
            print(f"[warn] Root '{parts[0]}' not found; fallback to Inbox.")
            folder = ns.GetDefaultFolder(6)
            start_idx = 0  # treat full path as under Inbox
        else:
            start_idx = 1

    # Descend into subfolders
    for name in parts[start_idx:]:
        found = None
        for i in range(1, folder.Folders.Count + 1):
            sub = folder.Folders.Item(i)
            if str(sub.Name).lower() == name.lower():
                found = sub
                break
        if not found:
            print(f"[warn] Subfolder '{name}' not found under '{folder.Name}', stop here.")
            break
        folder = found
    return folder
def _restrict_items(items, date_from: dt.datetime, date_to: Optional[dt.datetime], status: str):
    items.Sort("[ReceivedTime]", True)
    def fmt_ol(d: dt.datetime) -> str:
        return d.strftime("%d.%m.%Y %H:%M")
    clauses = [f"[ReceivedTime] >= '{fmt_ol(date_from)}'"]
    if date_to:
        eod = date_to.replace(hour=23, minute=59, second=59, microsecond=0)
        clauses.append(f"[ReceivedTime] <= '{fmt_ol(eod)}'")
    if status == "unread":
        clauses.append("[Unread] = True")
    elif status == "read":
        clauses.append("[Unread] = False")
    restriction = " AND ".join(clauses)
    print(f"[debug] Outlook Restrict: {restriction}")
    try:
        return items.Restrict(restriction)
    except Exception as e:
        print(f"[warn] Restrict failed ({e}); using unfiltered items.")
        return items

def _collect_from_items(restricted) -> List[EmailItem]:
    emails: List[EmailItem] = []
    for item in restricted:
        if getattr(item, "Class", None) != 43:  # olMail
            continue
        try:
            received = to_naive_local(item.ReceivedTime)
            subject  = str(item.Subject or "")
            sender_email = str(getattr(item, "SenderEmailAddress", "") or "")
            sender_name = str(getattr(item, "SenderName", "") or "")
            sender = f"{sender_name} <{sender_email}>"
            body_html  = str(getattr(item, "HTMLBody", "") or "")
            body_plain = str(getattr(item, "Body", "") or "")
            body_src   = body_html if len(body_html) > len(body_plain) else body_plain
            body_text  = html_to_text(body_src)
            to_recips  = str(getattr(item, "To", "") or "")
            cc_recips  = str(getattr(item, "CC", "") or "")
            conversation_id = str(getattr(item, "ConversationID", "") or "")
            entry_id = str(getattr(item, "EntryID", "") or "")
            sig = extract_signature(body_text)
            folder_path = str(getattr(getattr(item, "Parent", None), "FolderPath", "") or "") #DEBUG
            emails.append(EmailItem(
                received=received, subject=subject, sender=sender,
                to_recipients=to_recips, cc_recipients=cc_recips,
                body_text=body_text, conversation_id=conversation_id,
                entry_id=entry_id, signature_text=sig,
                folder_path=folder_path,  # DEBUG
            ))
        except Exception as e:
            print(f"[skip] Failed reading an item: {e}")
            continue
    return emails

def fetch_inbox_and_sent(date_from: dt.datetime,
                         date_to: Optional[dt.datetime],
                         status: str,
                         max_emails: int,
                         folder_path: str,
                         fetch_sent_too) -> List[EmailItem]:
    """Fetch Inbox + Sent Items, apply the same Restrict, then merge."""
    try:
        import win32com.client
    except ImportError:
        raise SystemExit("pywin32 is required. Install: pip install pywin32")
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    base = _resolve_folder(ns, folder_path)
    print(f"[i] Base folder resolved: {getattr(base, 'Name', '?')} ({getattr(base, 'FolderPath', '?')})")
    r_in = _restrict_items(base.Items, date_from, date_to, status)
    print(f"[i] Base after Restrict: {getattr(r_in, 'Count', '?')}")
    base_emails = _collect_from_items(r_in)
    sent_emails: List[EmailItem] = []

    if fetch_sent_too:
        sent = ns.GetDefaultFolder(5)
        print(f"[i] Sent folder: {getattr(sent, 'FolderPath', '?')}")
        r_out = _restrict_items(sent.Items, date_from, date_to, status)
        print(f"[i] Sent after Restrict: {getattr(r_out, 'Count', '?')}")
        all_sent = _collect_from_items(r_out)

        # Keep only Sent that belong to conversations seen in the base folder
        conv_ids = {e.conversation_id for e in base_emails if e.conversation_id}
        if conv_ids:
            sent_emails = [e for e in all_sent if e.conversation_id and e.conversation_id in conv_ids]
        else:
            sent_emails = []  # no conversations in base - ignore Sent completely

        print(f"[i] Sent kept after conv filter: {len(sent_emails)}")
    emails = base_emails + sent_emails
    # Optional hard cap after merge
    if max_emails and len(emails) > max_emails:
        emails.sort(key=lambda x: x.received, reverse=True)
        emails = emails[:max_emails]
    return emails
