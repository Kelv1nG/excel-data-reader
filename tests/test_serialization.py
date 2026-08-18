from __future__ import annotations

import json
from datetime import date, datetime, time, timedelta
from decimal import Decimal

import pytest

from excel_data_reader import FormulaValue, to_json, to_jsonable


def test_json_contract_tags_non_json_excel_scalars() -> None:
    value = {
        "date": date(2026, 8, 19),
        "datetime": datetime(2026, 8, 19, 14, 30, 15),
        "time": time(14, 30, 15),
        "duration": timedelta(days=1, seconds=30),
        "decimal": Decimal("12.340"),
        "bytes": b"xlsx",
        "nan": float("nan"),
    }

    converted = to_jsonable(value)

    assert converted["date"] == {"$type": "date", "value": "2026-08-19"}
    assert converted["datetime"] == {
        "$type": "datetime",
        "value": "2026-08-19T14:30:15",
    }
    assert converted["time"] == {"$type": "time", "value": "14:30:15"}
    assert converted["duration"] == {"$type": "timedelta", "value": 86430.0}
    assert converted["decimal"] == {"$type": "decimal", "value": "12.340"}
    assert converted["bytes"] == {
        "$type": "bytes",
        "encoding": "base64",
        "value": "eGxzeA==",
    }
    assert converted["nan"] == {"$type": "float", "value": "nan"}


def test_json_contract_serializes_public_dataclasses() -> None:
    serialized = json.loads(to_json(FormulaValue("=A1*2", 10), indent=2))

    assert serialized == {"formula": "=A1*2", "cached": 10}


def test_json_contract_rejects_unknown_values() -> None:
    with pytest.raises(TypeError, match="not supported"):
        to_jsonable(object())
