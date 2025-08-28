
import re
import json
import datetime as dt
from typing import Optional, List, Dict, Any
import os, sys
from bs4 import BeautifulSoup

def html_to_text(html: str) -> str:
    """Convert HTML to plain text with basic cleanup."""
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style"]):
        tag.decompose()
    return soup.get_text("\n", strip=True)

SIGNATURE_CUES = [
    "-- ",
    "S pozdravem", "S pozdravy", "S úctou",
    "Best regards", "Kind regards", "Regards",
    "Mit freundlichen Grüßen", "С уважением",
]

def extract_signature(body_text: str, max_sig_lines: int = 18) -> str:
    """Return probable signature block based on cues or tail heuristics."""
    if not body_text:
        return ""
    txt = body_text.strip()
    lower = txt.lower()
    positions = [lower.rfind(cue.lower()) for cue in SIGNATURE_CUES]
    pos = max(p for p in positions if p != -1) if any(p != -1 for p in positions) else -1
    if pos != -1:
        sig = txt[pos:]
    else:
        lines = [ln.rstrip() for ln in txt.splitlines()]
        tail = lines[-max_sig_lines:]
        rx_contact = re.compile(r"(@|\b(tel|phone|mobil|mobile|gsm|email|e-mail|www|web|fax)\b|\+\d|https?://)", re.I)
        picked = [ln for ln in tail if ln and (len(ln) <= 120) and (rx_contact.search(ln) or "," in ln or "|" in ln)]
        sig = "\n".join(picked[-max_sig_lines:]) if picked else "\n".join(tail)
    if len(sig) > 2000:
        sig = sig[:2000]
    return sig.strip()

def coerce_json(text: str) -> Optional[Dict[str, Any]]:
    """Parse JSON from model output, tolerant to code fences / noise."""
    if not text:
        return None
    m = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, flags=re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass
    m = re.search(r'(\{.*\})', text, flags=re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass
    try:
        return json.loads(text.replace("'", '"'))
    except Exception:
        return None

def coerce_to_schema(obj: dict, headers: List[str]) -> dict:
    """Ensure object has exactly the schema keys (strings only)."""
    clean = {}
    for h in headers:
        v = obj.get(h, "")
        if v is None:
            v = ""
        elif not isinstance(v, str):
            v = str(v)
        clean[h] = v
    return clean

def to_naive_local(d: dt.datetime) -> dt.datetime:
    """Convert aware datetime to local tz then drop tzinfo. Leave naive untouched."""
    if d is None:
        return None
    if d.tzinfo is None:
        return d
    try:
        return d.astimezone().replace(tzinfo=None)
    except Exception:
        return d.replace(tzinfo=None)

def is_incoming_email(item, my_emails: set) -> bool:
    """Return True if the message is incoming, False if outgoing."""
    sender_l = (item.sender or "").lower()
    return not (sender_l and sender_l in my_emails)
def _app_dir() -> str:
    """Return directory of the running app (exe in frozen, source in dev)."""
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))

def resolve_template_path(p: str) -> str:
    """Find template near exe or in bundled resources. Raise if missing."""
    if not p:
        raise FileNotFoundError("Empty TEMPLATE_XLSX path")
        # Keep absolute path if it exists
    if os.path.isabs(p) and os.path.exists(p):
            return os.path.normpath(p)

        # Same folder as the executable/script
    candidate = os.path.join(_app_dir(), os.path.basename(p))
    if os.path.exists(candidate):
            return os.path.normpath(candidate)
    raise FileNotFoundError(f"Template not found near EXE: {candidate}")

