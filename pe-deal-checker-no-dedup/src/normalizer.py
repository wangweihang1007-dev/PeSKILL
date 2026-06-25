"""Project short-name normalization utilities."""

from __future__ import annotations

import re
import unicodedata


def normalize_project_short_name(value: object) -> str:
    """Normalize a project short name for deterministic string matching."""
    if value is None:
        return ""

    text = unicodedata.normalize("NFKC", str(value))
    text = text.strip().lower()
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", text)
    return text

