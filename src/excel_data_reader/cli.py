"""Command-line inspection and explainable table discovery."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path

from excel_data_reader.diagnostics import ExcelDataReaderError
from excel_data_reader.model import (
    BodyPolicy,
    DiscoveryReport,
    TableQuery,
    WorkbookInventory,
)
from excel_data_reader.reader import ExcelReader
from excel_data_reader.serialization import to_json


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-data-reader",
        description="Inspect XLSX-family workbooks and explain table discovery.",
    )
    commands = parser.add_subparsers(dest="command", required=True)

    inspect = commands.add_parser("inspect", help="list workbook-authored structure")
    inspect.add_argument("workbook", type=Path)
    inspect.add_argument("--json", action="store_true", dest="as_json")

    find = commands.add_parser("find", help="find tables by normalized header names")
    find.add_argument("workbook", type=Path)
    find.add_argument(
        "--headers",
        required=True,
        help="comma-separated required logical headers",
    )
    find.add_argument("--optional", default="", help="comma-separated optional headers")
    find.add_argument(
        "--alias",
        action="append",
        default=[],
        metavar="FIELD=ALIAS|ALIAS",
        help="repeatable aliases for one declared field",
    )
    find.add_argument("--sheet")
    find.add_argument("--within", help="finite A1 search rectangle, such as A1:M5000")
    find.add_argument("--near", help="A1 cell used to resolve repeated matches")
    find.add_argument(
        "--adjacent-only",
        action="store_true",
        help="reject matches whose selected columns are separated",
    )
    find.add_argument(
        "--body",
        choices=("blank-rows", "last-populated", "explicit"),
        default="blank-rows",
    )
    find.add_argument("--blank-rows", type=int, default=2)
    find.add_argument("--bottom-row", type=int)
    find.add_argument("--json", action="store_true", dest="as_json")
    return parser


def _csv_headers(value: str) -> tuple[str, ...]:
    return tuple(item.strip() for item in value.split(",") if item.strip())


def _aliases(values: Sequence[str]) -> Mapping[str, tuple[str, ...]]:
    parsed: dict[str, list[str]] = {}
    for value in values:
        field, separator, raw_aliases = value.partition("=")
        aliases = [item.strip() for item in raw_aliases.split("|") if item.strip()]
        if not separator or not field.strip() or not aliases:
            raise ValueError(f"invalid alias {value!r}; expected FIELD=ALIAS|ALIAS")
        parsed.setdefault(field.strip(), []).extend(aliases)
    return {field: tuple(aliases) for field, aliases in parsed.items()}


def _body_policy(args: argparse.Namespace) -> BodyPolicy:
    if args.body == "blank-rows":
        if args.bottom_row is not None:
            raise ValueError("--bottom-row requires --body explicit")
        return BodyPolicy.until_blank_rows(args.blank_rows)
    if args.body == "last-populated":
        if args.bottom_row is not None:
            raise ValueError("--bottom-row requires --body explicit")
        return BodyPolicy.last_populated()
    if args.bottom_row is None:
        raise ValueError("--body explicit requires --bottom-row")
    return BodyPolicy.through_row(args.bottom_row)


def _query(args: argparse.Namespace) -> TableQuery:
    return TableQuery(
        required_headers=_csv_headers(args.headers),
        optional_headers=_csv_headers(args.optional),
        aliases=_aliases(args.alias),
        sheet=args.sheet,
        allow_non_adjacent_columns=not args.adjacent_only,
        body=_body_policy(args),
        near=args.near,
        within=args.within,
    )


def _print_inventory(path: Path, inventory: WorkbookInventory) -> None:
    print(f"Workbook: {path}")
    print(f"Sheets ({len(inventory.sheets)}):")
    for sheet in inventory.sheets:
        bounds = "empty" if sheet.apparent_bounds is None else sheet.apparent_bounds.a1
        tables = ", ".join(sheet.table_names) or "none"
        print(f"  {sheet.name}: state={sheet.state}, bounds={bounds}, tables={tables}")
    print(f"Native tables ({len(inventory.native_tables)}):")
    for table in inventory.native_tables:
        print(f"  {table.sheet}!{table.bounds.a1}: {table.name}")
    print(f"Named ranges ({len(inventory.named_ranges)}):")
    for named in inventory.named_ranges:
        destinations = ", ".join(item.a1 for item in named.destinations) or "unresolved"
        print(f"  {named.name}: {destinations}")


def _print_report(path: Path, report: DiscoveryReport) -> None:
    print(f"Workbook: {path}")
    print("Scans:")
    for scan in report.scans:
        bounds = "none" if scan.bounds is None else scan.bounds.a1
        state = "complete" if scan.completed else "incomplete"
        print(f"  {scan.sheet}!{bounds}: {scan.cells_considered} cells, {state}")
        for diagnostic in scan.diagnostics:
            print(f"    {diagnostic}")

    print("Candidates:")
    if not report.candidates:
        print("  none")
    for candidate in report.candidates:
        state = "SELECTED" if candidate.selected else "REJECTED"
        bounds = "unknown" if candidate.bounds is None else candidate.bounds.a1
        label = candidate.name or candidate.source.value
        matched = [item.requested_header for item in candidate.evidence if item.matched]
        missing = [
            item.requested_header
            for item in candidate.evidence
            if item.required and not item.matched
        ]
        reasons = ",".join(reason.value for reason in candidate.reasons) or "none"
        print(
            f"  {state} {candidate.sheet}!{bounds} ({label}): "
            f"matched={matched}, missing={missing}, reasons={reasons}"
        )

    print(f"Matches ({len(report.selected_matches)}):")
    for match in report.selected_matches:
        logical = [column.requested_header or column.name for column in match.columns]
        print(f"  {match.sheet}!{match.range}: source={match.source.value}, columns={logical}")
    for diagnostic in report.diagnostics:
        print(f"Diagnostic: {diagnostic}")


def _inspect(args: argparse.Namespace) -> int:
    with ExcelReader.open(args.workbook) as reader:
        inventory = reader.inventory()
    if args.as_json:
        print(to_json(inventory, indent=2))
    else:
        _print_inventory(args.workbook, inventory)
    return 0


def _find(args: argparse.Namespace) -> int:
    query = _query(args)
    with ExcelReader.open(args.workbook) as reader:
        report = reader.explain(query)
    if args.as_json:
        print(to_json(report, indent=2))
    else:
        _print_report(args.workbook, report)
    if len(report.selected_matches) == 1:
        return 0
    return 2 if not report.selected_matches else 3


def main(argv: Sequence[str] | None = None) -> int:
    """Run the command-line interface and return a process-compatible exit code."""

    parser = _parser()
    args = parser.parse_args(argv)
    try:
        return _inspect(args) if args.command == "inspect" else _find(args)
    except (ExcelDataReaderError, OSError, ValueError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
