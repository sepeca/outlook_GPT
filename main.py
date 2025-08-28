#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Entry point.
- Load config
- Fetch Outlook emails
- Post-filter by date and cap
- Build conversations
- Keep latest OUT message per conversation (configurable)
- Call LLM with single prompt schema (contact + inline summary)
- Export to Excel/template
"""
import sys, io
import os
import datetime as dt
from collections import defaultdict
from typing import List, Dict

import pandas as pd
from dotenv import load_dotenv

from config import (
    OUTPUT_DIR, OUTPUT_NAME, STRICT_SCHEMA,
    TEMPLATE_XLSX, TEMPLATE_SHEET, TEMPLATE_START_AT_R3,
    DATE_FROM_ENV, DATE_TO_ENV, DAYS_BACK_DEFAULT, MAX_EMAILS_DEFAULT,
    OUTLOOK_FOLDER_DEFAULT, STATUS_DEFAULT, MY_EMAILS, ENV_FILE, FETCH_SENT_TOO
)
from models import EmailItem
from utils import to_naive_local, coerce_to_schema, is_incoming_email, resolve_template_path
from outlook_io import fetch_inbox_and_sent
from template_export import export_rows_to_template
from gpt_client import call_gpt_with_prompts
from prompts import SCHEMA_KEYS_OSOBA, make_prompts_for_message

def _force_utf8_stdio():
    # Force UTF-8 for both streams. Safe in frozen and non-frozen modes.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        sys.stdout = io.TextIOWrapper(getattr(sys.stdout, "buffer", sys.stdout), encoding="utf-8", errors="replace")
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        sys.stderr = io.TextIOWrapper(getattr(sys.stderr, "buffer", sys.stderr), encoding="utf-8", errors="replace")

_force_utf8_stdio()
load_dotenv()

def main():
    # Output path
    _ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    _final_name = f"{OUTPUT_NAME or 'outlook_analysis'}_{_ts}.xlsx"
    output = os.path.normpath(os.path.join(OUTPUT_DIR or ".", _final_name))
    os.makedirs(os.path.dirname(output) or ".", exist_ok=True)

    # Date range
    date_from = to_naive_local(dt.datetime.fromisoformat(DATE_FROM_ENV)) if DATE_FROM_ENV else None
    date_to   = to_naive_local(dt.datetime.fromisoformat(DATE_TO_ENV))   if DATE_TO_ENV   else None
    if not date_from:
        date_from = to_naive_local(dt.datetime.now() - dt.timedelta(days=DAYS_BACK_DEFAULT))
    if date_to and date_to < date_from:
        df_from = date_to.replace(hour=0, minute=0, second=0, microsecond=0)
        df_to = date_from.replace(hour=23, minute=59, second=59, microsecond=0)
    else:
        df_from = date_from.replace(hour=0, minute=0, second=0, microsecond=0)
        df_to   = date_to.replace(hour=23, minute=59, second=59, microsecond=0) if date_to else None

    print(f"[i] Fetching emails... sources=Inbox+Sent range={[df_from, df_to]} status={STATUS_DEFAULT}")
    emails = fetch_inbox_and_sent(
        date_from=df_from, date_to=df_to,
        status=STATUS_DEFAULT, max_emails=MAX_EMAILS_DEFAULT,
        folder_path=OUTLOOK_FOLDER_DEFAULT,
        fetch_sent_too=FETCH_SENT_TOO
    )

    # Normalize datetimes
    for em in emails:
        em.received = to_naive_local(em.received)

    # In-memory post-filter guard
    before = len(emails)
    emails = [em for em in emails if (em.received >= df_from) and (df_to is None or em.received <= df_to)]
    print(f"[i] After date filter: {len(emails)} emails. Removed: {before - len(emails)}")

    if not emails:
        print("[i] Nothing to do.")
        return

    # Optional cap AFTER filtering
    if len(emails) > MAX_EMAILS_DEFAULT:
        emails.sort(key=lambda x: x.received, reverse=True)
        emails = emails[:MAX_EMAILS_DEFAULT]
        print(f"[i] Capped to MAX_EMAILS={MAX_EMAILS_DEFAULT}")


    # Build conversations
    conv_map: Dict[str, List[EmailItem]] = defaultdict(list)
    for em in emails:
        em.is_incoming = is_incoming_email(em, MY_EMAILS)
        conv_id = em.conversation_id or f"__{em.entry_id or id(em)}"
        conv_map[conv_id].append(em)

    # Select last message per conversation
    last_emails: List[EmailItem] = []
    for conv_id, lst in conv_map.items():
        if not lst:
            continue
        lst.sort(key=lambda x: x.received)  # ascending
        last_emails.append(lst[-1])



    # Send to GPT and collect rows
    print(f"[i] Conversations selected (latest-only): {len(last_emails)}")
    if not last_emails:
        print("[i] No conversations match selection. Done.")
        return

    # Send to GPT and collect rows
    rows: List[dict] = []
    total = len(last_emails)
    for idx, em in enumerate(last_emails, 1):
        try:
            conv_id = em.conversation_id or f"__{em.entry_id or id(em)}"
            # --- NEW: pre-call log per item ---
            print(
                f"Pokrok v praci na dopisech {idx}/{total}"
            )

            msg = {
                "is_incoming": bool(em.is_incoming),
                "received": em.received.strftime("%Y-%m-%d %H:%M"),
                "sender": em.sender,
                "to": em.to_recipients,
                "cc": em.cc_recipients,
                "subject": em.subject,
                "body": em.body_text[:20000],
                "signature": em.signature_text,
            }
            system_prompt, user_prompt = make_prompts_for_message(msg,[])
            obj = call_gpt_with_prompts(system_prompt, user_prompt)

            row = coerce_to_schema(obj or {}, SCHEMA_KEYS_OSOBA) if STRICT_SCHEMA else (obj or {})
            row["_EMAIL_RECEIVED"] = em.received.strftime("%Y-%m-%d %H:%M")
            row["_EMAIL_FROM"] = em.sender
            row["_EMAIL_SUBJECT"] = em.subject
            row["_EMAIL_DIR"] = ("IN" if em.is_incoming else "OUT")
            row["_CONV_ID"] = conv_id
            row["_SIGNATURE"] = em.signature_text
            rows.append(row)
        except Exception as e:
            # --- keep failure visible in export ---
            fallback = {k: "" for k in SCHEMA_KEYS_OSOBA}
            fallback["_ERROR"] = str(e)
            fallback["_EMAIL_SUBJECT"] = em.subject
            fallback["_CONV_ID"] = conv_id
            rows.append(fallback)
            print(f"[gpt] !! failed on conv={conv_id}: {e}")

    # Export
    print("[i] Exporting...")
    if TEMPLATE_XLSX:
        try:
            template_path = resolve_template_path(TEMPLATE_XLSX)  # NEW
            start_row = 3 if TEMPLATE_START_AT_R3 else None
            export_rows_to_template(
                # template_path=TEMPLATE_XLSX,  # OLD
                template_path=template_path,  # NEW
                out_path=output,
                sheet_name=TEMPLATE_SHEET,
                rows=rows,
                start_row=start_row
            )
            print(f"[ok] Saved {len(rows)} row(s) into template: {output}")
        except FileNotFoundError as e:
            print(f"[err] {e}")  # ASCII-safe log
            return
    else:
        # Union columns respecting SCHEMA_KEYS_OSOBA order
        union_cols, seen = [], set()
        for k in SCHEMA_KEYS_OSOBA:
            if k not in seen:
                union_cols.append(k); seen.add(k)
        for r in rows:
            for k in r.keys():
                if k not in seen:
                    union_cols.append(k); seen.add(k)
        df = pd.DataFrame(rows, columns=union_cols)
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Analysis")
        print(f"[ok] Saved {len(rows)} row(s) to: {output}")

    print("[done]")

if __name__ == "__main__":
    main()
