"""Deterministic JSON conversion for public reader models and cell values."""

from __future__ import annotations

import base64
import json
import math
from collections.abc import Mapping
from dataclasses import fields, is_dataclass
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from enum import Enum
from pathlib import Path
from typing import Any

JSON_VALUE_SCHEMA_VERSION = 1


def to_jsonable(value: Any) -> Any:
    """Convert a public model or Excel scalar into JSON-compatible values.

    Scalars that JSON cannot represent without losing their type use a small
    tagged object with ``$type`` and ``value`` keys.

    Args:
        value: Public model, container, or Excel scalar to convert.
    """

    if value is None or isinstance(value, (str, int, bool)):
        return value
    if isinstance(value, float):
        if math.isfinite(value):
            return value
        label = "nan" if math.isnan(value) else ("infinity" if value > 0 else "-infinity")
        return {"$type": "float", "value": label}
    if isinstance(value, Enum):
        return to_jsonable(value.value)
    if isinstance(value, datetime):
        return {"$type": "datetime", "value": value.isoformat()}
    if isinstance(value, date):
        return {"$type": "date", "value": value.isoformat()}
    if isinstance(value, time):
        return {"$type": "time", "value": value.isoformat()}
    if isinstance(value, timedelta):
        return {"$type": "timedelta", "value": value.total_seconds()}
    if isinstance(value, Decimal):
        return {"$type": "decimal", "value": str(value)}
    if isinstance(value, bytes):
        encoded = base64.b64encode(value).decode("ascii")
        return {"$type": "bytes", "encoding": "base64", "value": encoded}
    if isinstance(value, Path):
        return str(value)
    if is_dataclass(value) and not isinstance(value, type):
        return {field.name: to_jsonable(getattr(value, field.name)) for field in fields(value)}
    if isinstance(value, Mapping):
        return {str(key): to_jsonable(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [to_jsonable(item) for item in value]
    raise TypeError(f"{type(value).__name__} is not supported by the JSON contract")


def to_json(value: Any, *, indent: int | None = None) -> str:
    """Serialize public values without non-standard JSON constants.

    Args:
        value: Public model, container, or Excel scalar to serialize.
        indent: Number of spaces used for pretty-printing, or ``None`` for
            compact output.
    """

    return json.dumps(
        to_jsonable(value),
        indent=indent,
        ensure_ascii=False,
        allow_nan=False,
    )
