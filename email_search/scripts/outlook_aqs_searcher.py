#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook AQS-lite search over Excel exports.

Supported:
- from:, to:, cc:, bcc:, participants:
- subject:, body:
- hasattachment:, isflagged:
- received:, sent: (date or date range via ..)
- case:, site:, lpo:
- quoted phrases and OR/AND (top-level, no grouping)
"""

from __future__ import annotations

import argparse
import io
import json
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, Any

import difflib
import re
import pandas as pd


# Windows console encoding safety
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")


TRUE_VALUES = {"yes", "true", "1", "y", "t"}
FALSE_VALUES = {"no", "false", "0", "n", "f"}


@dataclass
class SchemaReport:
    sheet_name: str
    mapping: Dict[str, List[str]]
    missing_required: List[str]
    missing_optional: List[str]


class SchemaValidator:
    REQUIRED_COLUMNS = {
        "subject": ["Subject", "subject"],
        "from": ["SenderName", "SenderEmail", "From", "from", "Sender"],
        "to": ["RecipientTo", "To", "to", "Recipients"],
        "body": ["PlainTextBody", "Body", "body", "TextBody"],
        "delivery_time": ["DeliveryTime", "ReceivedTime", "SentOn", "Received"],
        "creation_time": ["CreationTime", "SentTime", "Created"],
    }

    OPTIONAL_COLUMNS = {
        "cc": ["RecipientCc", "Cc", "cc"],
        "bcc": ["RecipientBcc", "Bcc", "bcc"],
        "hasattachment": ["HasAttachment", "hasattachment"],
        "isflagged": ["IsFlagged", "isflagged"],
        "category": ["Category", "category"],
    }

    @staticmethod
    def _merge_aliases(primary: List[str], secondary: List[str]) -> List[str]:
        seen = set()
        merged: List[str] = []
        for alias in primary + secondary:
            if alias not in seen:
                merged.append(alias)
                seen.add(alias)
        return merged

    @classmethod
    def load_custom_aliases(
        cls, config_path: Optional[Path], file_name: str, sheet_name: str
    ) -> Dict[str, List[str]]:
        if not config_path or not config_path.exists():
            return {}
        try:
            config = json.loads(config_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

        file_config = config.get(file_name, {})
        if sheet_name in file_config:
            return file_config[sheet_name]
        if "default" in file_config:
            return file_config["default"]
        if "default" in config:
            return config["default"]
        return {}

    @classmethod
    def validate_and_map(
        cls,
        df: pd.DataFrame,
        sheet_name: str,
        custom_aliases: Optional[Dict[str, List[str]]] = None,
    ) -> SchemaReport:
        mapping: Dict[str, List[str]] = {}
        missing_required: List[str] = []
        missing_optional: List[str] = []

        required_columns = dict(cls.REQUIRED_COLUMNS)
        optional_columns = dict(cls.OPTIONAL_COLUMNS)

        if custom_aliases:
            for key, aliases in custom_aliases.items():
                if key in required_columns:
                    required_columns[key] = cls._merge_aliases(aliases, required_columns[key])
                elif key in optional_columns:
                    optional_columns[key] = cls._merge_aliases(aliases, optional_columns[key])
                else:
                    optional_columns[key] = aliases

        for canonical, aliases in required_columns.items():
            found = [alias for alias in aliases if alias in df.columns]
            if found:
                mapping[canonical] = found
            else:
                missing_required.append(canonical)

        for canonical, aliases in optional_columns.items():
            found = [alias for alias in aliases if alias in df.columns]
            if found:
                mapping[canonical] = found
            else:
                missing_optional.append(canonical)

        return SchemaReport(
            sheet_name=sheet_name,
            mapping=mapping,
            missing_required=missing_required,
            missing_optional=missing_optional,
        )


class EmailNormalizer:
    EXTRA_TEXT_COLUMNS = [
        "case_numbers",
        "hvdc_cases",
        "primary_case",
        "site",
        "sites",
        "primary_site",
        "lpo",
        "lpo_numbers",
        "category",
    ]

    @staticmethod
    def normalize_series(series: pd.Series) -> pd.Series:
        text = series.fillna("").astype(str)
        text = text.str.replace("_x000D_", " ", regex=False)
        text = text.str.replace(r"\s+", " ", regex=True)
        text = text.str.strip().str.lower()
        return text

    @classmethod
    def combine_series(cls, series_list: Iterable[pd.Series]) -> Optional[pd.Series]:
        series_list = [s for s in series_list if s is not None]
        if not series_list:
            return None
        combined = series_list[0]
        for series in series_list[1:]:
            combined = combined.str.cat(series, sep=" ", na_rep="")
        combined = combined.str.replace(r"\s+", " ", regex=True).str.strip()
        return combined

    @classmethod
    def create_normalized_columns(cls, df: pd.DataFrame, schema: SchemaReport) -> pd.DataFrame:
        df_norm = df.copy()
        mapping = schema.mapping

        def make_field(key: str, target: str) -> None:
            if key in mapping:
                series = cls.combine_series(
                    cls.normalize_series(df_norm[col]) for col in mapping[key]
                )
                if series is not None:
                    df_norm[target] = series

        make_field("subject", "__subject_lc")
        make_field("from", "__from_lc")
        make_field("to", "__to_lc")
        make_field("cc", "__cc_lc")
        make_field("bcc", "__bcc_lc")
        make_field("body", "__body_lc")
        make_field("category", "__category_lc")

        case_cols = [c for c in ["case_numbers", "hvdc_cases", "primary_case"] if c in df_norm.columns]
        if case_cols:
            df_norm["__case_lc"] = cls.combine_series(
                cls.normalize_series(df_norm[c]) for c in case_cols
            )

        site_cols = [c for c in ["site", "sites", "primary_site"] if c in df_norm.columns]
        if site_cols:
            df_norm["__site_lc"] = cls.combine_series(
                cls.normalize_series(df_norm[c]) for c in site_cols
            )

        lpo_cols = [c for c in ["lpo", "lpo_numbers"] if c in df_norm.columns]
        if lpo_cols:
            df_norm["__lpo_lc"] = cls.combine_series(
                cls.normalize_series(df_norm[c]) for c in lpo_cols
            )

        df_norm["__participants_lc"] = cls.combine_series(
            s for s in [
                df_norm.get("__from_lc"),
                df_norm.get("__to_lc"),
                df_norm.get("__cc_lc"),
                df_norm.get("__bcc_lc"),
            ]
            if s is not None
        )

        blob_parts = [
            df_norm.get("__subject_lc"),
            df_norm.get("__from_lc"),
            df_norm.get("__to_lc"),
            df_norm.get("__body_lc"),
            df_norm.get("__case_lc"),
            df_norm.get("__site_lc"),
            df_norm.get("__lpo_lc"),
            df_norm.get("__category_lc"),
        ]
        df_norm["__blob_lc"] = cls.combine_series(s for s in blob_parts if s is not None)

        if "delivery_time" in mapping:
            df_norm["__delivery_time"] = pd.to_datetime(
                df_norm[mapping["delivery_time"][0]], errors="coerce"
            )
        if "creation_time" in mapping:
            df_norm["__creation_time"] = pd.to_datetime(
                df_norm[mapping["creation_time"][0]], errors="coerce"
            )

        return df_norm


class AQSParser:
    KEYWORD_ALIASES = {
        "from": "from",
        "sender": "from",
        "to": "to",
        "cc": "cc",
        "bcc": "bcc",
        "participants": "participants",
        "subject": "subject",
        "body": "body",
        "hasattachment": "hasattachment",
        "isflagged": "isflagged",
        "category": "category",
        "received": "received",
        "sent": "sent",
        "case": "case",
        "site": "site",
        "lpo": "lpo",
    }

    @staticmethod
    def tokenize(query: str) -> List[str]:
        tokens: List[str] = []
        buf: List[str] = []
        in_quote = False
        quote_char = ""

        for ch in query:
            if in_quote:
                if ch == quote_char:
                    in_quote = False
                else:
                    buf.append(ch)
                continue

            if ch in ("'", '"'):
                in_quote = True
                quote_char = ch
                continue

            if ch in ("(", ")"):
                if buf:
                    tokens.append("".join(buf))
                    buf = []
                tokens.append(ch)
                continue

            if ch.isspace():
                if buf:
                    tokens.append("".join(buf))
                    buf = []
                continue

            buf.append(ch)

        if buf:
            tokens.append("".join(buf))

        return tokens

    @staticmethod
    def parse_date_range(expr: str) -> Tuple[Optional[pd.Timestamp], Optional[pd.Timestamp]]:
        expr = expr.strip()
        if ".." in expr:
            start_raw, end_raw = expr.split("..", 1)
            start = pd.to_datetime(start_raw.strip(), errors="coerce") if start_raw.strip() else None
            end = pd.to_datetime(end_raw.strip(), errors="coerce") if end_raw.strip() else None
            if end is not None and pd.notna(end):
                end = end + timedelta(days=1) - timedelta(seconds=1)
            return (start if pd.notna(start) else None, end if pd.notna(end) else None)

        single = pd.to_datetime(expr, errors="coerce")
        if pd.isna(single):
            return None, None
        start = single.replace(hour=0, minute=0, second=0)
        end = single.replace(hour=23, minute=59, second=59)
        return start, end

    @classmethod
    def parse_query(cls, query: str) -> Tuple[Dict[str, Any], List[str]]:
        tokens = cls.tokenize(query)
        stream = _TokenStream(tokens)
        warnings: List[str] = []

        ast = cls._parse_expression(stream, warnings)
        if stream.peek() is not None:
            remaining = stream.remaining()
            warnings.append(f"Unparsed tokens: {' '.join(remaining)}")

        return ast, warnings

    @classmethod
    def _parse_expression(cls, stream: "_TokenStream", warnings: List[str]) -> Dict[str, Any]:
        return cls._parse_or(stream, warnings)

    @classmethod
    def _parse_or(cls, stream: "_TokenStream", warnings: List[str]) -> Dict[str, Any]:
        left = cls._parse_and(stream, warnings)
        while True:
            token = stream.peek()
            if token is None or token == ")":
                break
            if token.upper() == "OR":
                stream.next()
                right = cls._parse_and(stream, warnings)
                left = {"kind": "or", "left": left, "right": right}
                continue
            break
        return left

    @classmethod
    def _parse_and(cls, stream: "_TokenStream", warnings: List[str]) -> Dict[str, Any]:
        left = cls._parse_not(stream, warnings)
        while True:
            token = stream.peek()
            if token is None or token == ")" or token.upper() == "OR":
                break
            if token.upper() == "AND":
                stream.next()
            right = cls._parse_not(stream, warnings)
            left = {"kind": "and", "left": left, "right": right}
        return left

    @classmethod
    def _parse_not(cls, stream: "_TokenStream", warnings: List[str]) -> Dict[str, Any]:
        token = stream.peek()
        if token is not None and token.upper() == "NOT":
            stream.next()
            operand = cls._parse_not(stream, warnings)
            return {"kind": "not", "operand": operand}
        return cls._parse_term(stream, warnings)

    @classmethod
    def _parse_term(cls, stream: "_TokenStream", warnings: List[str]) -> Dict[str, Any]:
        token = stream.next()
        if token is None:
            return {"kind": "empty"}

        if token == "(":
            expr = cls._parse_expression(stream, warnings)
            if stream.peek() == ")":
                stream.next()
            else:
                warnings.append("Unmatched opening parenthesis.")
            return expr

        if token == ")":
            warnings.append("Unmatched closing parenthesis.")
            return {"kind": "empty"}

        if ":" in token:
            field_raw, value_raw = token.split(":", 1)
            field = cls.KEYWORD_ALIASES.get(field_raw.lower())
            value = value_raw.strip()
            if not field:
                return {"kind": "text", "value": token}
            if not value:
                warnings.append(f"Empty value for field '{field_raw}'.")
                return {"kind": "empty"}

            if field in {"received", "sent"}:
                start, end = cls.parse_date_range(value)
                if not start and not end:
                    warnings.append(f"Invalid date range for '{field_raw}': {value}")
                    return {"kind": "empty"}
                return {"kind": "date", "field": field, "start": start, "end": end}

            if field in {"hasattachment", "isflagged"}:
                val = value.lower()
                if val in TRUE_VALUES:
                    return {"kind": "flag", "field": field, "value": True}
                if val in FALSE_VALUES:
                    return {"kind": "flag", "field": field, "value": False}
                warnings.append(f"Invalid flag value for '{field_raw}': {value}")
                return {"kind": "empty"}

            return {"kind": "field", "field": field, "value": value}

        return {"kind": "text", "value": token}


class _TokenStream:
    def __init__(self, tokens: List[str]) -> None:
        self._tokens = tokens
        self._pos = 0

    def peek(self) -> Optional[str]:
        if self._pos >= len(self._tokens):
            return None
        return self._tokens[self._pos]

    def next(self) -> Optional[str]:
        if self._pos >= len(self._tokens):
            return None
        token = self._tokens[self._pos]
        self._pos += 1
        return token

    def remaining(self) -> List[str]:
        return self._tokens[self._pos :]


class OutlookAqsSearcher:
    FIELD_TO_COLUMN = {
        "subject": "__subject_lc",
        "from": "__from_lc",
        "to": "__to_lc",
        "cc": "__cc_lc",
        "bcc": "__bcc_lc",
        "participants": "__participants_lc",
        "body": "__body_lc",
        "case": "__case_lc",
        "site": "__site_lc",
        "lpo": "__lpo_lc",
        "category": "__category_lc",
    }

    def __init__(
        self,
        excel_path: str,
        sheet: Optional[str] = None,
        auto_normalize: bool = True,
        config_path: Optional[str] = None,
    ):
        self.excel_path = Path(excel_path)
        self.sheet = sheet
        self.schema: Optional[SchemaReport] = None
        self.df: Optional[pd.DataFrame] = None
        self.df_normalized: Optional[pd.DataFrame] = None
        self.last_warnings: List[str] = []
        self.alias_source_map: Dict[str, str] = {}
        self.custom_aliases_used: Optional[Dict[str, List[str]]] = None
        self.config_path: Optional[str] = config_path

        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel file not found: {self.excel_path}")

        excel_file = pd.ExcelFile(self.excel_path)
        sheet_name = sheet or excel_file.sheet_names[0]
        if sheet_name not in excel_file.sheet_names:
            raise ValueError(f"Sheet not found: {sheet_name}")

        self.df = pd.read_excel(excel_file, sheet_name=sheet_name)
        custom_aliases = None
        if config_path:
            custom_aliases = SchemaValidator.load_custom_aliases(
                Path(config_path), self.excel_path.name, sheet_name
            )
            if custom_aliases:
                self.custom_aliases_used = custom_aliases
        self.schema = SchemaValidator.validate_and_map(
            self.df, sheet_name, custom_aliases=custom_aliases
        )

        if auto_normalize:
            self.df_normalized = EmailNormalizer.create_normalized_columns(self.df, self.schema)

        # Load synonyms
        self.synonyms = {}
        synonyms_path = self.excel_path.parent.parent / "config" / "synonyms.json"
        if synonyms_path.exists():
            try:
                raw_synonyms = json.loads(synonyms_path.read_text(encoding="utf-8"))
                for category in raw_synonyms.values():
                    for key, values in category.items():
                        # Normalize key
                        k_norm = key.lower()
                        if k_norm not in self.synonyms:
                            self.synonyms[k_norm] = set()
                        self.synonyms[k_norm].add(key)
                        for v in values:
                            self.synonyms[k_norm].add(v)
                            v_norm = v.lower()
                            if v_norm not in self.synonyms:
                                self.synonyms[v_norm] = set()
                            self.synonyms[v_norm].add(key)
                            for v2 in values:
                                if v != v2:
                                    self.synonyms[v_norm].add(v2)
            except Exception as e:
                print(f"Warning: Failed to load synonyms: {e}")

        if self.schema and self.df is not None:
            for canonical in self.schema.mapping.keys():
                if custom_aliases and canonical in custom_aliases:
                    config_found = any(
                        alias in self.df.columns for alias in custom_aliases[canonical]
                    )
                    builtin_aliases = SchemaValidator.REQUIRED_COLUMNS.get(
                        canonical, SchemaValidator.OPTIONAL_COLUMNS.get(canonical, [])
                    )
                    builtin_found = any(alias in self.df.columns for alias in builtin_aliases)
                    if config_found and builtin_found:
                        self.alias_source_map[canonical] = "merged"
                    elif config_found:
                        self.alias_source_map[canonical] = "config"
                    else:
                        self.alias_source_map[canonical] = "built-in"
                else:
                    self.alias_source_map[canonical] = "built-in"

    def _apply_text_search(self, series: pd.Series, value: str, fuzzy: bool = False) -> pd.Series:
        val_lower = value.lower()
        if not fuzzy:
            return series.str.contains(val_lower, na=False, regex=False)
        
        # Fuzzy match: check if 'value' is close to any token in the text
        def check_fuzzy(text):
            if pd.isna(text):
                return False
            text_str = str(text)
            # Fast path: exact substring
            if val_lower in text_str:
                return True
            
            # Tokenize and check similarity
            # This is expensive, so we use it sparingly
            tokens = text_str.split()
            matches = difflib.get_close_matches(val_lower, tokens, n=1, cutoff=0.8)
            return bool(matches)

        return series.apply(check_fuzzy)

    def _get_base_mask(self, token: dict, warnings: List[str], fuzzy: bool = False) -> pd.Series:
        if self.df_normalized is None:
            raise ValueError("Normalized data is not available.")

        if token["kind"] == "text":
            blob = self.df_normalized.get("__blob_lc")
            if blob is None:
                warnings.append("No searchable text columns found.")
                return pd.Series([False] * len(self.df_normalized))
            return self._apply_text_search(blob, token["value"], fuzzy=fuzzy)

        if token["kind"] == "field":
            field = token["field"]
            column = self.FIELD_TO_COLUMN.get(field)
            if not column or column not in self.df_normalized:
                warnings.append(f"Missing column for field '{field}'.")
                return pd.Series([False] * len(self.df_normalized))
            return self._apply_text_search(self.df_normalized[column], token["value"], fuzzy=fuzzy)

        if token["kind"] == "flag":
            field = token["field"]
            if not self.schema or field not in self.schema.mapping:
                warnings.append(f"Missing column for flag '{field}'.")
                return pd.Series([False] * len(self.df_normalized))
            col = self.schema.mapping[field][0]
            series = self.df[col]
            if series.dtype == bool:
                return series == token["value"]
            lowered = series.fillna("").astype(str).str.lower()
            is_true = lowered.isin(TRUE_VALUES)
            return is_true if token["value"] else ~is_true
        
        if token["kind"] == "date":
            field = token["field"]
            column = "__delivery_time" if field == "received" else "__creation_time"
            if column not in self.df_normalized:
                warnings.append(f"Missing date column for '{field}'.")
                return pd.Series([False] * len(self.df_normalized))
            series = self.df_normalized[column]
            start = token.get("start")
            end = token.get("end")
            if start and end:
                return (series >= start) & (series <= end)
            if start:
                return series >= start
            if end:
                return series <= end
            return pd.Series([False] * len(self.df_normalized))

        return pd.Series([False] * len(self.df_normalized))

    def _eval_ast(self, node: Dict[str, Any], warnings: List[str], fuzzy: bool = False) -> pd.Series:
        if self.df_normalized is None:
            raise ValueError("Normalized data is not available.")

        kind = node.get("kind")
        if kind == "empty":
            return pd.Series([False] * len(self.df_normalized))
        if kind == "and":
            return self._eval_ast(node["left"], warnings, fuzzy) & self._eval_ast(node["right"], warnings, fuzzy)
        if kind == "or":
            return self._eval_ast(node["left"], warnings, fuzzy) | self._eval_ast(node["right"], warnings, fuzzy)
        if kind == "not":
            return ~self._eval_ast(node["operand"], warnings, fuzzy)
        return self._get_base_mask(node, warnings, fuzzy=fuzzy)

    def expand_query(self, query: str) -> str:
        """
        Expand query terms with synonyms.
        Simple heuristic: if a token matches a synonym key, expand it to (token OR synonym1 OR synonym2).
        Only expands free text terms, not field:value pairs.
        """
        if not self.synonyms:
            return query

        tokens = AQSParser.tokenize(query)
        expanded_tokens = []
        
        for token in tokens:
            # Skip field:value, parens, operators
            if ":" in token or token in ("(", ")", "OR", "AND", "NOT"):
                expanded_tokens.append(token)
                continue
            
            # Check for synonyms
            token_lower = token.lower()
            if token_lower in self.synonyms:
                syns = self.synonyms[token_lower]
                # Create OR group
                # Quote multi-word synonyms
                group = []
                for s in syns:
                    if " " in s:
                        group.append(f'"{s}"')
                    else:
                        group.append(s)
                
                # Add original token if not in syns (it should be, but safety first)
                if token not in group and f'"{token}"' not in group:
                     group.append(token)
                     
                expanded_tokens.append(f"({' OR '.join(group)})")
            else:
                expanded_tokens.append(token)
                
        return " ".join(expanded_tokens)

    def calculate_relevance(self, df: pd.DataFrame, query: str) -> pd.DataFrame:
        if df.empty:
            return df
        
        # 1. Tokenize and Expand
        # Extract simple terms for scoring (ignore fields for now)
        raw_tokens = [t for t in query.split() if ":" not in t and "(" not in t and ")" not in t]
        if not raw_tokens:
            return df

        # Expand synonyms for scoring
        scoring_terms = set()
        for t in raw_tokens:
            t_lower = t.lower()
            scoring_terms.add(t_lower)
            if t_lower in self.synonyms:
                for s in self.synonyms[t_lower]:
                    scoring_terms.add(s.lower())
        
        scores = pd.Series(0.0, index=df.index)
        
        # Weights
        W_SUBJECT_EXACT = 20.0
        W_SUBJECT = 10.0
        W_SENDER = 5.0
        W_BODY = 1.0
        
        # Regex for entities (Project Tags, BLs)
        # HVDC-..., 4 letters + 7+ digits
        ENTITY_PATTERN = re.compile(r"(hvdc-[a-z0-9-]+|[a-z]{4}[0-9]{7,})", re.IGNORECASE)
        
        # Intent keywords (from Ontology)
        INTENTS = {"urgent", "action", "request", "fyi", "eta", "cost", "gate", "crane", "manifest", "risk"}

        for term in scoring_terms:
            # Check if term is an entity or intent
            is_entity = bool(ENTITY_PATTERN.match(term))
            is_intent = term in INTENTS
            
            term_weight_mult = 1.0
            if is_entity: term_weight_mult = 2.0
            if is_intent: term_weight_mult = 1.5

            # Subject
            if "__subject_lc" in df.columns:
                # Exact match bonus
                mask_exact = df["__subject_lc"] == term
                scores[mask_exact] += W_SUBJECT_EXACT * term_weight_mult
                
                # Contains
                mask = df["__subject_lc"].str.contains(re.escape(term), regex=True, na=False)
                scores[mask] += W_SUBJECT * term_weight_mult
            
            # Sender
            if "__from_lc" in df.columns:
                mask = df["__from_lc"].str.contains(re.escape(term), regex=True, na=False)
                scores[mask] += W_SENDER * term_weight_mult
                
            # Body
            if "__body_lc" in df.columns:
                mask = df["__body_lc"].str.contains(re.escape(term), regex=True, na=False)
                scores[mask] += W_BODY * term_weight_mult
                
        df = df.copy()
        df["_score"] = scores
        
        # Sort by score desc, then date desc
        sort_cols = ["_score"]
        ascending = [False]
        if "__delivery_time" in df.columns:
            sort_cols.append("__delivery_time")
            ascending.append(False)
            
        return df.sort_values(sort_cols, ascending=ascending)

    def search(self, query: str, max_results: int = 100, fuzzy: bool = False) -> pd.DataFrame:
        if self.df_normalized is None:
            raise ValueError("Normalized data is not available.")

        # Expand query with synonyms
        expanded_query = self.expand_query(query)
        
        ast, warnings = AQSParser.parse_query(expanded_query)
        self.last_warnings = warnings

        overall_mask = self._eval_ast(ast, self.last_warnings, fuzzy=fuzzy)
        results = self.df_normalized[overall_mask].copy()
        
        # Apply relevance scoring (use original query for scoring to prioritize exact matches)
        results = self.calculate_relevance(results, query)
        
        # Fallback sort if no score (should be handled by calculate_relevance but just in case)
        if "_score" not in results.columns and "__delivery_time" in results.columns:
            results = results.sort_values("__delivery_time", ascending=False)
            
        return results.head(max_results)

    def display_results(self, results_df: pd.DataFrame, show_body: bool = False, body_chars: int = 200) -> None:
        if results_df.empty:
            print("\n[No results]")
            return

        mapping = self.schema.mapping if self.schema else {}
        subject_col = (mapping.get("subject") or [None])[0]
        from_col = (mapping.get("from") or [None])[0]
        delivery_col = (mapping.get("delivery_time") or [None])[0]
        body_col = (mapping.get("body") or [None])[0]
        id_col = "no" if "no" in results_df.columns else None

        print(f"\n[Results] {len(results_df)}")
        print("=" * 80)
        for idx, (_, row) in enumerate(results_df.iterrows(), 1):
            label_id = f"{row.get(id_col, 'n/a')}" if id_col else "n/a"
            print(f"\n[{idx}] id: {label_id}")
            print(f"  subject: {row.get(subject_col, 'n/a')}")
            print(f"  from: {row.get(from_col, 'n/a')}")
            print(f"  received: {row.get(delivery_col, 'n/a')}")
            if show_body and body_col:
                body = str(row.get(body_col, ""))[:body_chars]
                print(f"  body: {body}...")

    def export_schema_report(self, output_file: Path, include_resolved: bool = True) -> None:
        if not self.schema:
            raise ValueError("Schema not available.")

        resolved_aliases: Dict[str, Any] = {}
        if include_resolved:
            for canonical, found_columns in self.schema.mapping.items():
                all_aliases: List[str] = []
                if canonical in SchemaValidator.REQUIRED_COLUMNS:
                    all_aliases.extend(SchemaValidator.REQUIRED_COLUMNS[canonical])
                elif canonical in SchemaValidator.OPTIONAL_COLUMNS:
                    all_aliases.extend(SchemaValidator.OPTIONAL_COLUMNS[canonical])

                if self.custom_aliases_used and canonical in self.custom_aliases_used:
                    custom_aliases = self.custom_aliases_used[canonical]
                    all_aliases = list(dict.fromkeys(custom_aliases + all_aliases))

                resolved_aliases[canonical] = {
                    "resolved_columns": found_columns,
                    "first_match": found_columns[0] if found_columns else None,
                    "source": self.alias_source_map.get(canonical, "built-in"),
                    "all_aliases_tried": all_aliases,
                    "matched_aliases": found_columns,
                    "unmatched_aliases": [alias for alias in all_aliases if alias not in found_columns],
                }

        report = {
            "file_name": str(self.excel_path.resolve()),
            "file_name_only": self.excel_path.name,
            "generated_at": datetime.now().isoformat(),
            "sheet_name": self.schema.sheet_name,
            "total_rows": int(len(self.df)) if self.df is not None else 0,
            "total_columns": int(len(self.df.columns)) if self.df is not None else 0,
            "mapping": self.schema.mapping,
            "resolved_aliases": resolved_aliases if include_resolved else None,
            "missing_required": self.schema.missing_required,
            "missing_optional": self.schema.missing_optional,
            "available_columns": list(self.df.columns) if self.df is not None else [],
            "config_file_used": str(Path(self.config_path).resolve())
            if self.config_path
            else None,
            "custom_aliases_applied": self.custom_aliases_used if self.custom_aliases_used else None,
        }
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")


def write_run_report(report_path: Path, report: dict) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Outlook AQS-lite email search")
    parser.add_argument("excel_file", help="Excel file path")
    parser.add_argument("--sheet", help="Sheet name (default: first sheet)")
    parser.add_argument(
        "--query", "-q", required=True, help="AQS query (supports parentheses and NOT)"
    )
    parser.add_argument("--max-results", "-n", type=int, default=50, help="Max results (default: 50)")
    parser.add_argument("--fuzzy", action="store_true", help="Enable fuzzy matching (slower)")
    parser.add_argument("--show-body", action="store_true", help="Show body preview")
    parser.add_argument("--export", help="Export results (.xlsx or .csv)")
    parser.add_argument("--schema-report", help="Write schema report JSON with resolved aliases")
    parser.add_argument("--report", help="Write run report JSON")
    parser.add_argument("--config", help="Path to column aliases config JSON file")
    parser.add_argument(
        "--auto-schema",
        action="store_true",
        help="Auto-generate schema report: <excel_file>.schema.json",
    )

    args = parser.parse_args()

    start_time = datetime.now()
    searcher = OutlookAqsSearcher(
        args.excel_file,
        sheet=args.sheet,
        auto_normalize=True,
        config_path=args.config,
    )

    if args.auto_schema:
        auto_report_path = Path(args.excel_file).with_suffix(".schema.json")
        searcher.export_schema_report(auto_report_path, include_resolved=True)
        print(f"\n[Auto Schema Report] {auto_report_path}")

    if args.schema_report:
        searcher.export_schema_report(Path(args.schema_report), include_resolved=True)
        print(f"\n[Schema Report] {args.schema_report}")

    results = searcher.search(args.query, max_results=args.max_results, fuzzy=args.fuzzy)
    elapsed = (datetime.now() - start_time).total_seconds()

    searcher.display_results(results, show_body=args.show_body)
    if searcher.last_warnings:
        print("\n[Warnings]")
        for warning in searcher.last_warnings:
            print(f"- {warning}")
    print(f"\n[Elapsed] {elapsed:.3f}s")

    if args.export:
        export_path = Path(args.export)
        export_path.parent.mkdir(parents=True, exist_ok=True)
        if export_path.suffix.lower() == ".csv":
            results.to_csv(export_path, index=False)
        else:
            results.to_excel(export_path, index=False, engine="openpyxl")
        print(f"\n[Exported] {export_path}")

    if args.report:
        report = {
            "task": "aqs_search",
            "input_file": str(Path(args.excel_file).resolve()),
            "sheet": searcher.schema.sheet_name if searcher.schema else None,
            "query": args.query,
            "results_count": int(len(results)),
            "elapsed_seconds": round(elapsed, 3),
            "warnings": searcher.last_warnings,
            "config": str(Path(args.config).resolve()) if args.config else None,
            "generated_at": datetime.now().isoformat(),
        }
        write_run_report(Path(args.report), report)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
