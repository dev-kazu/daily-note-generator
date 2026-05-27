from __future__ import annotations

import argparse
import sys
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .workbook import append_note, build_entry, validate_append_target


DEFAULT_FILE = Path("daily_notes.xlsx")
DEFAULT_TIMEZONE = "Asia/Tokyo"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m daily_note",
        description="Append a daily note to an Excel workbook.",
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=DEFAULT_FILE,
        help="Excel workbook path. Defaults to daily_notes.xlsx.",
    )
    parser.add_argument(
        "--timezone",
        default=DEFAULT_TIMEZONE,
        help="IANA timezone name. Defaults to Asia/Tokyo.",
    )
    parser.add_argument(
        "--note",
        help="Note text. If omitted, you will be prompted for one line of text.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be appended without writing the workbook.",
    )
    return parser


def _read_note_from_prompt() -> str:
    return input("Note: ")


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        timezone = ZoneInfo(args.timezone)
    except ZoneInfoNotFoundError:
        print(f"Error: unknown timezone: {args.timezone}", file=sys.stderr)
        return 2

    note = args.note if args.note is not None else _read_note_from_prompt()
    if not note.strip():
        print("Error: note must not be empty.", file=sys.stderr)
        return 1

    try:
        if args.dry_run:
            entry = build_entry(note, timezone)
            validate_append_target(args.file)
            print(
                f"Dry run: would add note to {args.file} "
                f"at {entry.date} {entry.time}."
            )
        else:
            entry = append_note(args.file, note, timezone)
            print(f"Added note to {args.file} at {entry.date} {entry.time}.")
    except OSError as error:
        print(f"Error: failed to write workbook: {error}", file=sys.stderr)
        return 1
    except ValueError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    return 0
