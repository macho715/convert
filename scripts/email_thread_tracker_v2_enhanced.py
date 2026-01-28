#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Heuristic threading with inverted indexes for Outlook Excel exports.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd

from email_derived_fields import EmailDerivedFields, extract_emails


@dataclass
class ThreadRelation:
    parent_idx: Optional[int] = None
    relation_type: str = "heuristic"
    confidence: float = 0.0
    thread_id: Optional[str] = None


class EmailThreadTrackerV2Enhanced:
    def __init__(self, df: pd.DataFrame) -> None:
        self.df = EmailDerivedFields.add_derived_fields(df)
        self.threads: Dict[str, List[int]] = defaultdict(list)
        self.thread_relations: Dict[int, ThreadRelation] = {}

        self.by_thread_key: Dict[str, Set[int]] = defaultdict(set)
        self.by_subject: Dict[str, Set[int]] = defaultdict(set)
        self.by_participants: Dict[str, Set[int]] = defaultdict(set)
        self.by_body_hash: Dict[str, Set[int]] = defaultdict(set)
        self.by_case: Dict[str, Set[int]] = defaultdict(set)
        self.by_site: Dict[str, Set[int]] = defaultdict(set)
        self.by_lpo: Dict[str, Set[int]] = defaultdict(set)
        self.by_sender: Dict[str, Set[int]] = defaultdict(set)
        self.by_time_window: Dict[str, Set[int]] = defaultdict(set)

        self._build_inverted_indexes()
        self._build_thread_tree()

    @staticmethod
    def _split_tokens(value: object) -> List[str]:
        if value is None or pd.isna(value):
            return []
        text = str(value)
        parts = [p.strip().upper() for p in text.replace(";", ",").split(",")]
        return [p for p in parts if p]

    def _build_inverted_indexes(self) -> None:
        for idx, row in self.df.iterrows():
            thread_key = row.get("thread_key_heuristic", "")
            if thread_key:
                self.by_thread_key[thread_key].add(idx)

            subject_norm = row.get("subject_norm", "")
            if subject_norm:
                self.by_subject[subject_norm].add(idx)

            participants = row.get("participants_norm", "")
            if participants:
                self.by_participants[participants].add(idx)

            body_hash = row.get("body_hash", "")
            if body_hash:
                self.by_body_hash[body_hash].add(idx)

            for col in ["case_numbers", "hvdc_cases", "primary_case", "entity_cases"]:
                for case in self._split_tokens(row.get(col, "")):
                    self.by_case[case].add(idx)

            for col in ["site", "sites", "primary_site", "entity_sites"]:
                for site in self._split_tokens(row.get(col, "")):
                    self.by_site[site].add(idx)

            for col in ["lpo", "lpo_numbers", "entity_lpos"]:
                for lpo in self._split_tokens(row.get(col, "")):
                    self.by_lpo[lpo].add(idx)

            sender_emails = extract_emails(row.get("SenderEmail")) + extract_emails(
                row.get("SenderName")
            )
            for email in sender_emails:
                self.by_sender[email.upper()].add(idx)

            if pd.notna(row.get("DeliveryTime")):
                try:
                    dt = pd.to_datetime(row.get("DeliveryTime"))
                    time_key = dt.strftime("%Y-%m-%d")
                    self.by_time_window[time_key].add(idx)
                except Exception:
                    pass

    def _build_thread_tree(self) -> None:
        thread_id_counter = 0
        processed: Set[int] = set()

        for idx, row in self.df.iterrows():
            if idx in processed:
                continue

            candidates = set()

            thread_key = row.get("thread_key_heuristic", "")
            if thread_key:
                candidates.update(self.by_thread_key.get(thread_key, set()))

            subject_norm = row.get("subject_norm", "")
            if subject_norm:
                candidates.update(self.by_subject.get(subject_norm, set()))

            participants = row.get("participants_norm", "")
            if participants:
                candidates.update(self.by_participants.get(participants, set()))

            body_hash = row.get("body_hash", "")
            if body_hash:
                candidates.update(self.by_body_hash.get(body_hash, set()))

            for col in ["case_numbers", "hvdc_cases", "primary_case", "entity_cases"]:
                for case in self._split_tokens(row.get(col, "")):
                    candidates.update(self.by_case.get(case, set()))

            for col in ["site", "sites", "primary_site", "entity_sites"]:
                for site in self._split_tokens(row.get(col, "")):
                    candidates.update(self.by_site.get(site, set()))

            for col in ["lpo", "lpo_numbers", "entity_lpos"]:
                for lpo in self._split_tokens(row.get(col, "")):
                    candidates.update(self.by_lpo.get(lpo, set()))

            sender_emails = extract_emails(row.get("SenderEmail")) + extract_emails(
                row.get("SenderName")
            )
            if sender_emails and pd.notna(row.get("DeliveryTime")):
                try:
                    dt = pd.to_datetime(row.get("DeliveryTime"))
                    for offset in range(-3, 4):
                        time_key = (dt + pd.Timedelta(days=offset)).strftime("%Y-%m-%d")
                        for email in sender_emails:
                            candidates.update(
                                self.by_sender.get(email.upper(), set())
                                & self.by_time_window.get(time_key, set())
                            )
                except Exception:
                    pass

            candidates.discard(idx)

            if not candidates:
                processed.add(idx)
                continue

            thread_id = f"thread_{thread_id_counter}"
            thread_id_counter += 1
            thread_indices = [idx] + sorted(candidates)
            self.threads[thread_id] = thread_indices

            for t_idx in thread_indices:
                if t_idx in processed:
                    continue
                confidence = self._calculate_thread_confidence(
                    idx, t_idx, thread_key, subject_norm, participants
                )
                self.thread_relations[t_idx] = ThreadRelation(
                    thread_id=thread_id,
                    relation_type="heuristic",
                    confidence=confidence,
                )
                processed.add(t_idx)
                self.df.at[t_idx, "confidence_thread"] = confidence
                self.df.at[t_idx, "thread_id"] = thread_id
                self.df.at[t_idx, "thread_size"] = len(thread_indices)

    def _calculate_thread_confidence(
        self,
        idx1: int,
        idx2: int,
        thread_key: str,
        subject_norm: str,
        participants: str,
    ) -> float:
        row1 = self.df.loc[idx1]
        row2 = self.df.loc[idx2]

        if thread_key and row2.get("thread_key_heuristic") == thread_key:
            return 0.85

        subject_match = subject_norm and row2.get("subject_norm") == subject_norm
        participants_match = participants and row2.get("participants_norm") == participants

        entity_match = False
        for col in ["case_numbers", "hvdc_cases", "entity_cases"]:
            val1 = str(row1.get(col, "")).upper()
            val2 = str(row2.get(col, "")).upper()
            if val1 and val2 and (val1 in val2 or val2 in val1):
                entity_match = True
                break

        time_proximity = 0.0
        try:
            dt1 = pd.to_datetime(row1.get("DeliveryTime"))
            dt2 = pd.to_datetime(row2.get("DeliveryTime"))
            if pd.notna(dt1) and pd.notna(dt2):
                days_diff = abs((dt1 - dt2).days)
                if days_diff <= 7:
                    time_proximity = 1.0 - (days_diff / 7.0)
        except Exception:
            pass

        confidence = 0.0
        if subject_match:
            confidence += 0.4
        if participants_match:
            confidence += 0.3
        if entity_match:
            confidence += 0.2
        if time_proximity > 0.5:
            confidence += 0.1 * time_proximity
        return min(confidence, 0.9)

    def search_with_context(self, query: str, max_results: int = 50) -> Tuple[pd.DataFrame, Dict]:
        query_upper = query.upper()
        matching_indices = set()

        for idx, row in self.df.iterrows():
            subject = str(row.get("Subject", "")).upper()
            body = str(row.get("PlainTextBody", "")).upper()
            sender = str(row.get("SenderName", "")).upper()
            if query_upper in subject or query_upper in body or query_upper in sender:
                matching_indices.add(idx)

        all_related = set(matching_indices)

        from email_derived_fields import extract_entities

        entities = extract_entities(query)
        for case in entities["cases"]:
            all_related.update(self.by_case.get(case, set()))
        for site in entities["sites"]:
            all_related.update(self.by_site.get(site, set()))
        for lpo in entities["lpos"]:
            all_related.update(self.by_lpo.get(lpo, set()))

        for idx in matching_indices:
            relation = self.thread_relations.get(idx)
            if relation and relation.thread_id in self.threads:
                all_related.update(self.threads[relation.thread_id])

        result_df = self.df.loc[list(all_related)].copy()
        if "DeliveryTime" in result_df.columns:
            result_df["DeliveryTime"] = pd.to_datetime(result_df["DeliveryTime"], errors="coerce")
            result_df = result_df.sort_values("DeliveryTime", ascending=False)

        confidences = [
            self.thread_relations[idx].confidence
            for idx in all_related
            if idx in self.thread_relations
        ]
        avg_confidence = sum(confidences) / max(len(confidences), 1)

        context = {
            "total_found": len(matching_indices),
            "total_with_context": len(all_related),
            "threads_included": len(
                {
                    self.thread_relations[idx].thread_id
                    for idx in all_related
                    if idx in self.thread_relations and self.thread_relations[idx].thread_id
                }
            ),
            "entities_found": entities,
            "avg_confidence": round(avg_confidence, 3),
        }

        return result_df.head(max_results), context
