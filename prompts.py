from inspect import signature
# ---------- Schema for person extraction (incoming emails) ----------

from typing import List, Dict, Tuple
import os

_my_names  = [s.strip() for s in os.getenv("MY_NAME", "").split(",") if s.strip()]
_my_emails = [s.strip().lower() for s in os.getenv("MY_EMAILS", "").split(",") if s.strip()]
_rules_raw = os.getenv("PROMPT_RULES", "") or ""
_my_rules = _rules_raw.replace("\\n", "\n").replace('\\"', '"').strip()
# NEW: include client name to map into Excel header "Název Klienta"
SCHEMA_KEYS_OSOBA = [
    "NazevKlienta",
    "Prijmeni",
    "Jmeno",
    "TitulPred",
    "TitulZa",
    "Funkce",
    "Tel1",
    "Email",
    "WWW",
    "PoznamkaKOsobe",
]

# Render a fixed-order JSON skeleton that the model must follow
_schema_json_block_osoba = "{\n" + ",\n".join([f'  "{k}": ""' for k in SCHEMA_KEYS_OSOBA]) + "\n}"

# System prompt for extracting a single JSON object for a person
SYSTEM_PROMPT_OSOBA = f"""
You are an assistant that extracts structured company/client data from emails and returns it as a SINGLE JSON object.

Important rules:
- Keys must match exactly and appear in the same order as listed below.
- All values must be strings.
- If the information is not available, use an empty string "".
- Do not add or remove keys.
- Return only the JSON, without any explanations or extra text.

Identity rules (who is ME vs the CONTACT):
- Treat these identities as ME and NEVER output ME as the contact person:
  - jmena a prijmeni: {_my_names}
  - emails: {_my_emails}
- If the current message appears authored by ME (From matches ME, or signature matches ME), DO NOT extract ME. Extract the COUNTERPART instead:
  - Prefer a single non‑ME person found in headers (From/To/Cc) or in the signature/body.
  - If multiple candidates exist, pick the primary counterpart (the main recipient or the signer of the current message).
- Never put ME's email/phone into output fields. If only ME's data is found, leave fields empty.

Field notes:
{_my_rules}



The JSON structure to follow:
{_schema_json_block_osoba}
""".strip()

# User prompt template for incoming message (full metadata + body of THIS email only)
USER_PROMPT_TEMPLATE_INCOMING = """
EMAIL METADATA
- received: {received}
- from: {sender}
- to: {to}
- cc: {cc}
- subject: {subject}

EMAIL BODY
\"\"\"{body}\"\"\"

EMAIL SIGNATURE
\"\"\"{signature}\"\"\"

If a company name is present, put it into "NazevKlienta". Extract the rest according to SYSTEM_PROMPT.
""".strip()


def make_prompts_for_message(msg: Dict, thread_messages: List[Dict]) -> Tuple[str, str]:
    """
    Decide which prompt pair to use based on message direction.

    msg: dict with keys:
      - is_incoming: bool
      - received, sender, to, cc, subject, body: str
      - signature: str (optional; used in incoming mode)
    thread_messages: list of dicts for the whole dialog, used only in outgoing mode.

    Returns:
      (system_prompt: str, user_prompt: str)
    """
    if msg.get("is_incoming", True):
        system_prompt = SYSTEM_PROMPT_OSOBA
        user_prompt = USER_PROMPT_TEMPLATE_INCOMING.format(
            received=msg.get("received", ""),
            sender=msg.get("sender", ""),
            to=msg.get("to", ""),
            cc=msg.get("cc", ""),
            subject=msg.get("subject", ""),
            body=msg.get("body", ""),
            signature=msg.get("signature", ""),
        )



    return system_prompt, user_prompt


