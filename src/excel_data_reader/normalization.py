"""Deterministic header normalization."""

from __future__ import annotations

import re
import unicodedata
from typing import Any

_WHITESPACE = re.compile(r"\s+")
_SEPARATORS = str.maketrans({"_": " ", "-": " ", "\u00a0": " "})


def normalize_header(value: Any) -> str:
    """Return the exact canonical form used by MVP header matching."""

    text = unicodedata.normalize("NFKC", str(value)).translate(_SEPARATORS)
    return _WHITESPACE.sub(" ", text).strip().casefold()
