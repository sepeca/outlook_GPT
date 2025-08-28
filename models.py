
from dataclasses import dataclass
import datetime as dt
from typing import Optional

@dataclass
class EmailItem:
    received: dt.datetime
    subject: str
    sender: str
    to_recipients: str
    cc_recipients: str
    body_text: str
    conversation_id: Optional[str] = None
    entry_id: Optional[str] = None
    is_incoming: Optional[bool] = None
    signature_text: str = ""
    folder_path: str = ""
