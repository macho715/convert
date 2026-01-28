# -*- coding: utf-8 -*-
from __future__ import annotations

import re


EMAIL_RE = re.compile(r"([^@]+)@(.+)")


def mask_email(email: str) -> str:
    if not email:
        return ""
    match = EMAIL_RE.match(email.strip())
    if not match:
        return email
    name, domain = match.groups()
    if len(name) <= 2:
        masked = name[:1] + "*"
    else:
        masked = name[:2] + "*" * (len(name) - 2)
    return f"{masked}@{domain}"


def normalize_body_text(text) -> str:
    if text is None:
        return ""
    body = str(text)
    body = body.replace("_x000D_", "\n")
    body = body.replace("\r\n", "\n").replace("\r", "\n")
    return body
