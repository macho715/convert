#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Thread tracker v3 (Excel-only, heuristic).
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List, Optional, Set
import re
import difflib

import pandas as pd


EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
SUBJECT_TAG_RE = re.compile(r"^\[(.*?)\]\s*")
SUBJECT_PREFIX_RE = re.compile(
    r"^(re(\[\d+\])?|fw|fwd|recall|reminder|회수|회신|답장|전달)\s*:\s*",
    re.IGNORECASE,
)
SUBJECT_EMPTY = {"(제목 없음)", "제목 없음", "(no subject)", "no subject"}

SITE_CODES = {"AGI", "DAS", "MIR", "MIRFA", "GHALLAN"}

CASE_PATTERNS = [
    re.compile(r"\bHVDC-[A-Z0-9-]+\b", re.IGNORECASE),
    re.compile(r"\bSCT-[A-Z0-9-]+\b", re.IGNORECASE),
    re.compile(r"\bJ71[-_]?\d+\b", re.IGNORECASE),
]

LPO_PATTERN = re.compile(r"\bLPO[-_ ]?\d+\b", re.IGNORECASE)

BLOCKED_SUBJECTS = {"", "RE", "FW", "FWD", "RECALL", "REMINDER"}


def _clean_text(value: Optional[str]) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("_x000D_", " ")
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_emails(value: Optional[str]) -> List[str]:
    cleaned = _clean_text(value)
    if not cleaned:
        return []
    return list(dict.fromkeys([m.lower() for m in EMAIL_RE.findall(cleaned)]))


def normalize_subject(subject: Optional[str]) -> str:
    cleaned = _clean_text(subject)
    if not cleaned:
        return ""
    if cleaned.strip().lower() in SUBJECT_EMPTY:
        return ""

    while True:
        match = SUBJECT_TAG_RE.match(cleaned)
        if not match:
            break
        cleaned = cleaned[match.end() :].lstrip()

    while True:
        match = SUBJECT_PREFIX_RE.match(cleaned)
        if not match:
            break
        cleaned = cleaned[match.end() :].lstrip()

    cleaned = cleaned.strip(" -:")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.upper()


def normalize_participants(
    sender_name: Optional[str],
    sender_email: Optional[str],
    recipient_to: Optional[str],
    recipient_cc: Optional[str] = None,
    recipient_bcc: Optional[str] = None,
) -> str:
    participants: Set[str] = set()

    sender_emails = extract_emails(sender_email) + extract_emails(sender_name)
    for email in sender_emails:
        participants.add(email)

    if not sender_emails and sender_name:
        label = _clean_text(sender_name).lower()
        if label:
            participants.add(f"name:{label}")

    for value in [recipient_to, recipient_cc, recipient_bcc]:
        for email in extract_emails(value):
            participants.add(email)

    return "|".join(sorted(participants))


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


@dataclass
class ThreadMeta:
    members: Set[int]
    start_dt: Optional[pd.Timestamp]
    end_dt: Optional[pd.Timestamp]
    subject_norm: str
    cases: List[str]
    sites: List[str]
    lpos: List[str]
    confidence: float


class EmailThreadTrackerV3:
    def __init__(self, df: pd.DataFrame) -> None:
        self.df = df.copy()
        self.thread_meta: Dict[str, ThreadMeta] = {}
        self.thread_by_row: Dict[int, str] = {}
        self.by_case: Dict[str, Set[int]] = defaultdict(set)
        self.by_site: Dict[str, Set[int]] = defaultdict(set)
        self.by_lpo: Dict[str, Set[int]] = defaultdict(set)
        self._build_derived_fields()
        self._build_indexes()
        self._build_threads()

    def _build_derived_fields(self) -> None:
        if "Subject" in self.df.columns:
            self.df["_subject_norm"] = self.df["Subject"].apply(normalize_subject)
        else:
            self.df["_subject_norm"] = ""

        self.df["_participants_norm"] = self.df.apply(
            lambda row: normalize_participants(
                row.get("SenderName"),
                row.get("SenderEmail"),
                row.get("RecipientTo"),
                row.get("RecipientCc"),
                row.get("RecipientBcc"),
            ),
            axis=1,
        )

        combined = self.df.get("Subject", "").fillna("").astype(str) + " " + self.df.get(
            "PlainTextBody", ""
        ).fillna("").astype(str)
        entities = combined.apply(extract_entities)
        self.df["_entity_cases"] = entities.apply(lambda x: ",".join(x["cases"]))
        self.df["_entity_sites"] = entities.apply(lambda x: ",".join(x["sites"]))
        self.df["_entity_lpos"] = entities.apply(lambda x: ",".join(x["lpos"]))

    @staticmethod
    def _split_tokens(raw: object) -> List[str]:
        if raw is None or pd.isna(raw):
            return []
        text = str(raw)
        if text.strip().lower() in {"nan", "none", "null"}:
            return []
        parts = [part.strip() for part in text.split(",")]
        return [p for p in parts if p and p.lower() not in {"nan", "none", "null"}]

    def _build_indexes(self) -> None:
        for idx, row in self.df.iterrows():
            for col in ["case_numbers", "hvdc_cases", "primary_case", "_entity_cases"]:
                for val in self._split_tokens(row.get(col, "")):
                    self.by_case[val.strip().upper()].add(idx)

            for col in ["site", "sites", "primary_site", "_entity_sites"]:
                for val in self._split_tokens(row.get(col, "")):
                    self.by_site[val.strip().upper()].add(idx)

            for col in ["lpo", "lpo_numbers", "_entity_lpos"]:
                for val in self._split_tokens(row.get(col, "")):
                    self.by_lpo[val.strip().upper()].add(idx)

    def _build_threads(self) -> None:
        n = len(self.df)
        parent = list(range(n))
        rank = [0] * n

        def find(x: int) -> int:
            if parent[x] != x:
                parent[x] = find(parent[x])
            return parent[x]

        def union(a: int, b: int) -> None:
            ra, rb = find(a), find(b)
            if ra == rb:
                return
            if rank[ra] < rank[rb]:
                parent[ra] = rb
            elif rank[ra] > rank[rb]:
                parent[rb] = ra
            else:
                parent[rb] = ra
                rank[ra] += 1

        buckets: Dict[str, List[int]] = defaultdict(list)

        def add_bucket_value(prefix: str, value: str, idx: int) -> None:
            if not value:
                return
            if value.strip().lower() in {"nan", "none", "null"}:
                return
            buckets[f"{prefix}:{value}"].append(idx)

        for idx, row in self.df.iterrows():
            subject_norm = str(row.get("_subject_norm", "")).strip().upper()
            participants_norm = str(row.get("_participants_norm", "")).strip()

            cases = {
                val.strip().upper()
                for col in ["case_numbers", "hvdc_cases", "primary_case", "_entity_cases"]
                for val in self._split_tokens(row.get(col, ""))
                if val.strip()
            }
            lpos = {
                val.strip().upper()
                for col in ["lpo", "lpo_numbers", "_entity_lpos"]
                for val in self._split_tokens(row.get(col, ""))
                if val.strip()
            }

            for case in cases:
                add_bucket_value("case", case, idx)
            for lpo in lpos:
                add_bucket_value("lpo", lpo, idx)

            if subject_norm and subject_norm not in BLOCKED_SUBJECTS:
                for case in cases:
                    add_bucket_value("subject_case", f"{subject_norm}||{case}", idx)
                for lpo in lpos:
                    add_bucket_value("subject_lpo", f"{subject_norm}||{lpo}", idx)

            if cases and lpos:
                for case in cases:
                    for lpo in lpos:
                        add_bucket_value("case_lpo", f"{case}||{lpo}", idx)

            if subject_norm and participants_norm and cases and subject_norm not in BLOCKED_SUBJECTS:
                for case in cases:
                    add_bucket_value(
                        "strong", f"{subject_norm}||{participants_norm}||{case}", idx
                    )

        for bucket_key, members in buckets.items():
            if len(members) < 2:
                continue
            if bucket_key.startswith("subject:") or bucket_key.startswith("participants:"):
                continue
            head = members[0]
            for member in members[1:]:
                union(head, member)

        threads: Dict[int, List[int]] = {}
        for idx in range(n):
            root = find(idx)
            threads.setdefault(root, []).append(idx)

        thread_id_counter = 0
        for members in threads.values():
            if len(members) < 2:
                continue
            if len(members) > 200:
                continue
            thread_id = f"thread_{thread_id_counter}"
            thread_id_counter += 1
            member_set = set(members)

            self.df.loc[members, "_thread_id"] = thread_id

            delivery = pd.to_datetime(
                self.df.loc[members].get("DeliveryTime", None), errors="coerce"
            )
            start_dt = delivery.min() if hasattr(delivery, "min") else None
            end_dt = delivery.max() if hasattr(delivery, "max") else None

            subject_norm = (
                self.df.loc[members, "_subject_norm"]
                .dropna()
                .astype(str)
                .mode()
                .iloc[0]
                if "_subject_norm" in self.df.columns
                else ""
            )

            cases = sorted(
                {
                    v.strip().upper()
                    for col in ["case_numbers", "hvdc_cases", "primary_case", "_entity_cases"]
                    if col in self.df.columns
                    for v in str(self.df.loc[members, col].dropna().astype(str).str.cat(sep=","))
                    .split(",")
                    if v.strip()
                }
            )
            sites = sorted(
                {
                    v.strip().upper()
                    for col in ["site", "sites", "primary_site", "_entity_sites"]
                    if col in self.df.columns
                    for v in str(self.df.loc[members, col].dropna().astype(str).str.cat(sep=","))
                    .split(",")
                    if v.strip()
                }
            )
            lpos = sorted(
                {
                    v.strip().upper()
                    for col in ["lpo", "lpo_numbers", "_entity_lpos"]
                    if col in self.df.columns
                    for v in str(self.df.loc[members, col].dropna().astype(str).str.cat(sep=","))
                    .split(",")
                    if v.strip()
                }
            )

            confidence = self._thread_confidence(members)

            self.thread_meta[thread_id] = ThreadMeta(
                members=member_set,
                start_dt=start_dt,
                end_dt=end_dt,
                subject_norm=subject_norm,
                cases=cases,
                sites=sites,
                lpos=lpos,
                confidence=confidence,
            )
            for idx in members:
                self.thread_by_row[idx] = thread_id

    def _thread_confidence(self, members: List[int]) -> float:
        if len(members) <= 1:
            return 0.0
        base = members[0]
        scores = [self._pair_confidence(base, other) for other in members[1:]]
        return round(sum(scores) / max(len(scores), 1), 4)

    def _pair_confidence(self, i: int, j: int) -> float:
        row1 = self.df.loc[i]
        row2 = self.df.loc[j]

        # 1. Subject Similarity (Weight: 0.45)
        subj1 = str(row1.get("_subject_norm", "")).upper()
        subj2 = str(row2.get("_subject_norm", "")).upper()
        
        sim_subject = 0.0
        if subj1 and subj2:
            # Use difflib for similarity ratio (0.0 to 1.0)
            sim_subject = difflib.SequenceMatcher(None, subj1, subj2).ratio()
            # Boost exact match slightly to handle short subjects better
            if subj1 == subj2:
                sim_subject = 1.0
        
        # 2. Entity Overlap (Weight: 0.25)
        # Collect entities from both rows
        ents1 = set()
        ents2 = set()
        
        for col in ["case_numbers", "hvdc_cases", "_entity_cases", "site", "sites", "_entity_sites", "lpo", "lpo_numbers", "_entity_lpos"]:
            v1 = str(row1.get(col, "")).upper()
            v2 = str(row2.get(col, "")).upper()
            if v1: ents1.update([x.strip() for x in v1.split(",") if x.strip()])
            if v2: ents2.update([x.strip() for x in v2.split(",") if x.strip()])
            
        sim_entity = 0.0
        if ents1 and ents2:
            intersection = len(ents1.intersection(ents2))
            union = len(ents1.union(ents2))
            if union > 0:
                sim_entity = intersection / union

        # 3. Time Decay (Weight: 0.15)
        time_decay = 0.0
        try:
            dt1 = pd.to_datetime(row1.get("DeliveryTime"))
            dt2 = pd.to_datetime(row2.get("DeliveryTime"))
            if pd.notna(dt1) and pd.notna(dt2):
                days_diff = abs((dt1 - dt2).days)
                if days_diff <= 14:
                    time_decay = 1.0 - (days_diff / 14.0)
                else:
                    time_decay = 0.0
        except Exception:
            pass

        # 4. Reply Hint (Weight: 0.15)
        hint_reply = 0.0
        # Check for RE/FW prefixes in original subject (before normalization)
        orig_subj1 = str(row1.get("Subject", "")).upper()
        orig_subj2 = str(row2.get("Subject", "")).upper()
        
        has_prefix = (
            orig_subj1.startswith("RE:") or orig_subj1.startswith("FW:") or
            orig_subj2.startswith("RE:") or orig_subj2.startswith("FW:")
        )
        
        # Check for In-Reply-To patterns in body (simple heuristic)
        # e.g., "From: ... Sent: ..." block
        body1 = str(row1.get("PlainTextBody", ""))
        body2 = str(row2.get("PlainTextBody", ""))
        has_quote = "From:" in body1 or "From:" in body2 or "Sent:" in body1 or "Sent:" in body2
        
        if has_prefix or has_quote:
            hint_reply = 1.0

        # Calculate Total Score
        # Weights: Subject(0.45) + Entity(0.25) + Time(0.15) + Reply(0.15)
        W_S, W_E, W_T, W_R = 0.45, 0.25, 0.15, 0.15
        
        score = (W_S * sim_subject) + (W_E * sim_entity) + (W_T * time_decay) + (W_R * hint_reply)
        
        return round(min(score, 1.0), 4)

    def get_pair_confidence(self, i: int, j: int) -> float:
        return self._pair_confidence(i, j)

    def export_threads(self) -> List[dict]:
        out: List[dict] = []
        for tid, meta in self.thread_meta.items():
            out.append(
                {
                    "thread_id": tid,
                    "members": sorted(meta.members),
                    "start_dt": meta.start_dt.isoformat() if meta.start_dt is not None else None,
                    "end_dt": meta.end_dt.isoformat() if meta.end_dt is not None else None,
                    "subject_norm": meta.subject_norm,
                    "cases": meta.cases,
                    "sites": meta.sites,
                    "lpos": meta.lpos,
                    "confidence": round(meta.confidence, 4),
                }
            )
        return out

    def search_with_context(self, query: str, max_results: int = 50) -> tuple[pd.DataFrame, dict]:
        query_upper = query.upper()
        matching_indices = set()

        for idx, row in self.df.iterrows():
            subject = str(row.get("Subject", "")).upper()
            body = str(row.get("PlainTextBody", "")).upper()
            sender = str(row.get("SenderName", "")).upper()
            if query_upper in subject or query_upper in body or query_upper in sender:
                matching_indices.add(idx)

        all_related = set(matching_indices)
        entities = extract_entities(query)

        for case in entities["cases"]:
            case_upper = case.upper()
            if self.by_case:
                all_related.update(self.by_case.get(case_upper, set()))
            else:
                for idx, row in self.df.iterrows():
                    val = str(row.get("case_numbers", "")).upper()
                    if case_upper in val:
                        all_related.add(idx)

        for site in entities["sites"]:
            site_upper = site.upper()
            if self.by_site:
                all_related.update(self.by_site.get(site_upper, set()))
            else:
                for idx, row in self.df.iterrows():
                    val = str(row.get("sites", "")).upper()
                    if site_upper in val:
                        all_related.add(idx)

        for lpo in entities["lpos"]:
            lpo_upper = lpo.upper()
            if self.by_lpo:
                all_related.update(self.by_lpo.get(lpo_upper, set()))
            else:
                for idx, row in self.df.iterrows():
                    val = str(row.get("lpo_numbers", "")).upper()
                    if lpo_upper in val:
                        all_related.add(idx)

        for idx in matching_indices:
            thread_id = self.thread_by_row.get(idx)
            if thread_id and thread_id in self.thread_meta:
                all_related.update(self.thread_meta[thread_id].members)

        result_df = self.df.loc[list(all_related)].copy()
        if "DeliveryTime" in result_df.columns:
            result_df["DeliveryTime"] = pd.to_datetime(result_df["DeliveryTime"], errors="coerce")
            result_df = result_df.sort_values("DeliveryTime", ascending=False)

        confidences = []
        for idx in all_related:
            thread_id = self.thread_by_row.get(idx)
            if thread_id and thread_id in self.thread_meta:
                confidences.append(self.thread_meta[thread_id].confidence)

        context = {
            "total_found": len(matching_indices),
            "total_with_context": len(all_related),
            "threads_included": len(
                {
                    self.thread_by_row.get(idx)
                    for idx in all_related
                    if self.thread_by_row.get(idx)
                }
            ),
            "entities_found": entities,
            "avg_confidence": round(sum(confidences) / max(len(confidences), 1), 4),
        }
        return result_df.head(max_results), context
