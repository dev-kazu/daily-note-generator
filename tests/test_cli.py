from __future__ import annotations

from openpyxl import load_workbook

from daily_note.cli import main


def test_cli_writes_note_from_argument(tmp_path):
    workbook_path = tmp_path / "notes.xlsx"

    exit_code = main(["--file", str(workbook_path), "--note", "hello"])

    workbook = load_workbook(workbook_path)
    worksheet = workbook["Notes"]

    assert exit_code == 0
    assert worksheet.max_row == 2
    assert worksheet.cell(row=2, column=4).value == "hello"


def test_cli_prompts_for_note_when_argument_is_missing(tmp_path, monkeypatch):
    workbook_path = tmp_path / "notes.xlsx"
    monkeypatch.setattr("builtins.input", lambda prompt: "prompted note")

    exit_code = main(["--file", str(workbook_path)])

    workbook = load_workbook(workbook_path)
    worksheet = workbook["Notes"]

    assert exit_code == 0
    assert worksheet.cell(row=2, column=4).value == "prompted note"


def test_cli_rejects_empty_note(tmp_path, capsys):
    workbook_path = tmp_path / "notes.xlsx"

    exit_code = main(["--file", str(workbook_path), "--note", " "])

    captured = capsys.readouterr()
    assert exit_code == 1
    assert "note must not be empty" in captured.err
    assert not workbook_path.exists()


def test_cli_rejects_unknown_timezone(tmp_path, capsys):
    workbook_path = tmp_path / "notes.xlsx"

    exit_code = main([
        "--file",
        str(workbook_path),
        "--timezone",
        "Invalid/Timezone",
        "--note",
        "hello",
    ])

    captured = capsys.readouterr()
    assert exit_code == 2
    assert "unknown timezone" in captured.err
    assert not workbook_path.exists()
