#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Derived fields for Outlook Excel exports tuned to OUTLOOK_HVDC_ALL_rev.xlsx.
"""

from __future__ import annotations

import hashlib
import re
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Set

import pandas as pd


EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
SUBJECT_TAG_RE = re.compile(r"^\[(.*?)\]\s*")
SUBJECT_PREFIX_RE = re.compile(
    r"^(re(\[\d+\])?|fw|fwd|recall|reminder|회수|회신|답장|전달)\s*:\s*",
    re.IGNORECASE,
)

SUBJECT_EMPTY = {
    "(제목 없음)",
    "제목 없음",
    "(no subject)",
    "no subject",
}

SITE_CODES = {"AGI", "DAS", "MIR", "MIRFA", "GHALLAN"}

CASE_PATTERNS = [
    re.compile(r"\bHVDC-[A-Z0-9-]+\b", re.IGNORECASE),
    re.compile(r"\bSCT-[A-Z0-9-]+\b", re.IGNORECASE),
    re.compile(r"\bJ71[-_]?\d+\b", re.IGNORECASE),
]

LPO_PATTERN = re.compile(r"\bLPO[-_ ]?\d+\b", re.IGNORECASE)


def _clean_text(value: Optional[str]) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("_x000D_", " ")
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_emails(text: Optional[str]) -> List[str]:
    cleaned = _clean_text(text)
    if not cleaned:
        return []
    emails = EMAIL_RE.findall(cleaned)
    return list(dict.fromkeys([e.lower() for e in emails]))


def normalize_subject(subject: Optional[str]) -> str:
    cleaned = _clean_text(subject)
    if not cleaned:
        return ""
    if cleaned.strip().lower() in SUBJECT_EMPTY:
        return ""

    # Remove leading bracket tags like [HVDC-AGI], [Approved], [EXTERNAL]
    while True:
        match = SUBJECT_TAG_RE.match(cleaned)
        if not match:
            break
        cleaned = cleaned[match.end() :].lstrip()

    # Remove prefix chains like RE:, FW:, Recall:
    while True:
        match = SUBJECT_PREFIX_RE.match(cleaned)
        if not match:
            break
        cleaned = cleaned[match.end() :].lstrip()

    cleaned = cleaned.strip(" -:")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.upper()


def parse_recipients(recipients_raw: Optional[str]) -> List[str]:
    return extract_emails(recipients_raw)


def normalize_participants(
    sender_name: Optional[str],
    sender_email: Optional[str],
    recipient_to: Optional[str],
    recipient_cc: Optional[str] = None,
    recipient_bcc: Optional[str] = None,
) -> str:
    participants: Set[str] = set()

    sender_emails = extract_emails(sender_email)
    sender_emails.extend(extract_emails(sender_name))
    for email in sender_emails:
        participants.add(email)

    if not sender_emails and sender_name:
        sender_label = _clean_text(sender_name).lower()
        if sender_label:
            participants.add(f"name:{sender_label}")

    for item in parse_recipients(recipient_to):
        participants.add(item)
    for item in parse_recipients(recipient_cc):
        participants.add(item)
    for item in parse_recipients(recipient_bcc):
        participants.add(item)

    return "|".join(sorted(participants))


def hash_body(body: Optional[str], length: int = 40) -> str:
    cleaned = _clean_text(body).lower()
    if not cleaned:
        return ""
    digest = hashlib.sha256(cleaned.encode("utf-8")).hexdigest()
    return digest[:length]


def create_thread_key_heuristic(
    subject_norm: str,
    participants_norm: str,
    delivery_time: Optional[datetime],
    time_window_days: int = 7,
) -> str:
    if delivery_time is not None and pd.notna(delivery_time):
        try:
            dt = pd.to_datetime(delivery_time)
            bucket_start = dt - timedelta(days=dt.weekday())
            time_bucket = bucket_start.strftime("%Y-W%U")
        except Exception:
            time_bucket = "unknown"
    else:
        time_bucket = "unknown"

    parts = [
        subject_norm[:120],
        participants_norm[:200],
        time_bucket,
    ]
    return "||".join(parts)


def extract_entities(text: Optional[str]) -> Dict[str, List[str]]:
    cleaned = _clean_text(text).upper()
    entities = {"cases": [], "sites": [], "lpos": []}
    if not cleaned:
        return entities

    cases: Set[str] = set()
    for pattern in CASE_PATTERNS:
        cases.update(pattern.findall(cleaned))
    entities["cases"] = sorted(cases)

    sites: Set[str] = set()
    for code in SITE_CODES:
        if re.search(rf"\b{re.escape(code)}\b", cleaned):
            sites.add(code)
    entities["sites"] = sorted(sites)

    lpos: Set[str] = set()
    for match in LPO_PATTERN.findall(cleaned):
        normalized = re.sub(r"[-_ ]", "", match.upper())
        if normalized.startswith("LPO"):
            normalized = normalized.replace("LPO", "LPO-", 1)
        lpos.add(normalized)
    entities["lpos"] = sorted(lpos)

    return entities


class EmailDerivedFields:
    @classmethod
    def add_derived_fields(cls, df: pd.DataFrame) -> pd.DataFrame:
        df_out = df.copy()

        if "Subject" in df_out.columns:
            df_out["subject_norm"] = df_out["Subject"].apply(normalize_subject)

        sender_name_col = "SenderName" if "SenderName" in df_out.columns else None
        sender_email_col = "SenderEmail" if "SenderEmail" in df_out.columns else None
        to_col = "RecipientTo" if "RecipientTo" in df_out.columns else None
        cc_col = "RecipientCc" if "RecipientCc" in df_out.columns else None
        bcc_col = "RecipientBcc" if "RecipientBcc" in df_out.columns else None

        if sender_name_col or sender_email_col or to_col:
            df_out["participants_norm"] = df_out.apply(
                lambda row: normalize_participants(
                    row.get(sender_name_col) if sender_name_col else None,
                    row.get(sender_email_col) if sender_email_col else None,
                    row.get(to_col) if to_col else None,
                    row.get(cc_col) if cc_col else None,
                    row.get(bcc_col) if bcc_col else None,
                ),
                axis=1,
            )

        if "PlainTextBody" in df_out.columns:
            df_out["body_hash"] = df_out["PlainTextBody"].apply(hash_body)

        delivery_col = "DeliveryTime" if "DeliveryTime" in df_out.columns else None
        if "subject_norm" in df_out.columns and "participants_norm" in df_out.columns:
            df_out["thread_key_heuristic"] = df_out.apply(
                lambda row: create_thread_key_heuristic(
                    row.get("subject_norm", ""),
                    row.get("participants_norm", ""),
                    pd.to_datetime(row.get(delivery_col), errors="coerce")
                    if delivery_col
                    else None,
                ),
                axis=1,
            )

        if "PlainTextBody" in df_out.columns or "Subject" in df_out.columns:
            subject_series = (
                df_out["Subject"] if "Subject" in df_out.columns else pd.Series("", index=df_out.index)
            )
            body_series = (
                df_out["PlainTextBody"]
                if "PlainTextBody" in df_out.columns
                else pd.Series("", index=df_out.index)
            )
            combined = subject_series.fillna("").astype(str) + " " + body_series.fillna("").astype(str)
            entities = combined.apply(extract_entities)
            df_out["entity_cases"] = entities.apply(lambda x: ",".join(x["cases"]))
            df_out["entity_sites"] = entities.apply(lambda x: ",".join(x["sites"]))
            df_out["entity_lpos"] = entities.apply(lambda x: ",".join(x["lpos"]))

        if "confidence_thread" not in df_out.columns:
            df_out["confidence_thread"] = 0.0

        return df_out
