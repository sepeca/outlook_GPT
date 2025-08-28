
from dotenv import load_dotenv
load_dotenv()
import os,sys
def _env_path():
    if getattr(sys, "frozen", False):
        base = os.getenv("APPDATA") or ""
        return os.path.join(base, "OutlookGPT", ".env")
    # dev path
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")

ENV_FILE = _env_path()
loaded = load_dotenv(ENV_FILE, override=True)

# DEBUG: show where we read .env from
try:
    print(f"[cfg] env={ENV_FILE} exists={os.path.exists(ENV_FILE)} loaded={loaded}")
except Exception:
    pass

def _clean(s: str) -> str:
    s = (s or "").strip().strip('"').strip("'")
    for ch in ("\u200b", "\u200e", "\u200f", "\u00a0"):
        s = s.replace(ch, "")
    return s

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini").strip()

DATE_FROM_ENV = os.getenv("DATE_FROM", "").strip()   # "YYYY-MM-DD" or empty
DATE_TO_ENV   = os.getenv("DATE_TO", "").strip()     # "YYYY-MM-DD" or empty

DAYS_BACK_DEFAULT      = int(os.getenv("DAYS_BACK", "7"))
MAX_EMAILS_DEFAULT     = int(os.getenv("MAX_EMAILS", "200"))
OUTLOOK_FOLDER_DEFAULT = os.getenv("OUTLOOK_FOLDER", "Inbox").strip()
STATUS_DEFAULT         = os.getenv("STATUS", "all").strip().lower()  # all|unread|read

OUTPUT_DIR  = os.getenv("OUTPUT_DIR", "").strip()
OUTPUT_NAME = os.getenv("OUTPUT_NAME", "outlook_analysis").strip()
STRICT_SCHEMA = os.getenv("STRICT_SCHEMA", "true").lower() == "true"

TEMPLATE_XLSX        = os.getenv("TEMPLATE_XLSX", "").strip()
TEMPLATE_SHEET       = os.getenv("TEMPLATE_SHEET", "").strip()
TEMPLATE_START_AT_R3 = os.getenv("TEMPLATE_START_AT_ROW3", "true").lower() == "true"

MY_EMAILS = {e.strip().lower() for e in os.getenv("MY_EMAILS", "").split(",") if e.strip()}

FETCH_SENT_TOO = os.getenv("FETCH_SENT_TOO", "true").lower() == "true"