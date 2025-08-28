
import os
import requests
from typing import Dict, Any

from config import OPENAI_API_KEY, OPENAI_BASE_URL, OPENAI_MODEL
from utils import coerce_json

# NEW: simple debug switch via env
DEBUG_GPT = os.getenv("DEBUG_GPT", "false").lower() == "true"


# NEW: monotonic counter for dumps
_req_counter = {"n": 0}


def _sprint(msg: str):
    try:
        print(msg)
    except Exception:
        # last resort: replace non-encodables
        print(msg.encode("utf-8", "replace").decode("utf-8"))
def call_gpt_with_prompts(system_prompt: str, user_prompt: str) -> Dict[str, Any]:
    """Send prepared prompts to LLM and return parsed JSON."""
    if not OPENAI_API_KEY:
        raise SystemExit("Set OPENAI_API_KEY in .env")

    _req_counter["n"] += 1
    # ---- PRE-LOG ----
    _sprint(f"[gpt] -> POST {OPENAI_BASE_URL.rstrip('/')}/chat/completions model={OPENAI_MODEL} req#{_req_counter['n']}")


    payload = {
        "model": OPENAI_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.0,
        "response_format": {"type": "json_object"}
    }
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    url = f"{OPENAI_BASE_URL.rstrip('/')}/chat/completions"

    try:
        r = requests.post(url, headers=headers, json=payload, timeout=90)
        r.raise_for_status()
        data = r.json()
        content = data["choices"][0]["message"]["content"]
        # ---- POST-LOG ----
        _sprint(f"[gpt] <- OK req#{_req_counter['n']} len={len(content)}")

        return coerce_json(content) or {}
    except requests.HTTPError as e:
        # Log server reply body for easier diagnosis
        body = getattr(e.response, "text", "") if hasattr(e, "response") else ""
        _sprint(f"[gpt] !! HTTP {getattr(e.response,'status_code',None)} req#{_req_counter['n']} body={body[:500]}")
        raise
    except Exception as e:
        _sprint(f"[gpt] !! ERROR req#{_req_counter['n']} {e}")
        raise
