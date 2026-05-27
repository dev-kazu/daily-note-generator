# daily-note-generator

Append short daily notes to a single Excel workbook.

## Setup

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -e ".[dev]"
```

## Usage

Prompt for a one-line note and append it to `daily_notes.xlsx`:

```bash
python -m daily_note
```

Append a note without the prompt:

```bash
python -m daily_note --note "Reviewed the release checklist"
```

Preview a note without writing to the workbook:

```bash
python -m daily_note --note "Reviewed the release checklist" --dry-run
```

Use a different workbook or timezone:

```bash
python -m daily_note --file data/daily_notes.xlsx --timezone Asia/Tokyo
```

The workbook uses a `Notes` sheet with these columns:

- `date`
- `time`
- `weekday`
- `note`

## Test

```bash
python -m pytest
```
