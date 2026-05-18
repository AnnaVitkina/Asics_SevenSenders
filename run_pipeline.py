#!/usr/bin/env python3
"""Interactive pipeline: clean selected input JSON files and export one Excel workbook."""

from __future__ import annotations

import sys
from pathlib import Path


def _find_project_root() -> Path:
    """Locate the folder containing json_clean.py (works with python and exec())."""
    try:
        script_dir = Path(__file__).resolve().parent
        if (script_dir / "json_clean.py").is_file():
            return script_dir
    except NameError:
        pass

    for candidate in (
        Path.cwd(),
        Path("/content/Asics_SevenSenders"),
        *Path.cwd().parents,
    ):
        if (candidate / "json_clean.py").is_file():
            return candidate.resolve()

    raise FileNotFoundError(
        "Could not find project root (json_clean.py). "
        "Run from the project folder or set PIPELINE_ROOT, e.g.\n"
        "  import os; os.chdir('/content/Asics_SevenSenders')"
    )


_root_override = globals().get("PIPELINE_ROOT")
PIPELINE_ROOT = (
    Path(_root_override).resolve() if _root_override else _find_project_root()
)
if str(PIPELINE_ROOT) not in sys.path:
    sys.path.insert(0, str(PIPELINE_ROOT))

from json_clean import process_file as clean_json_file
from main_costs_to_excel import transform_cleaned_jsons_to_xlsx

# ---------------------------------------------------------------------------
# Folder paths (defaults: input / processing / output under PIPELINE_ROOT).
# Override with absolute paths if needed, e.g. on Colab/Drive:
#   HARDCODED_INPUT_FOLDER = Path("/content/Asics_SevenSenders/input")
# ---------------------------------------------------------------------------
HARDCODED_INPUT_FOLDER = PIPELINE_ROOT / "input"
HARDCODED_PROCESSING_FOLDER = PIPELINE_ROOT / "processing"
HARDCODED_OUTPUT_FOLDER = PIPELINE_ROOT / "output"


def list_input_json_files() -> list[Path]:
    if not HARDCODED_INPUT_FOLDER.is_dir():
        raise FileNotFoundError(
            f"Input folder does not exist: {HARDCODED_INPUT_FOLDER}"
        )
    return sorted(HARDCODED_INPUT_FOLDER.glob("*.json"))


def prompt_file_selection(files: list[Path]) -> list[Path]:
    print(f"\nProject root: {PIPELINE_ROOT}")
    print(f"Input folder: {HARDCODED_INPUT_FOLDER}")
    print(f"Processing folder: {HARDCODED_PROCESSING_FOLDER}")
    print(f"Output folder: {HARDCODED_OUTPUT_FOLDER}\n")

    if not files:
        raise FileNotFoundError(
            f"No .json files found in {HARDCODED_INPUT_FOLDER}"
        )

    print("Available JSON files:")
    for idx, path in enumerate(files, 1):
        print(f"  {idx:2}. {path.name}")

    print(
        "\nEnter file numbers to process (e.g. 1 or 1,3,5), "
        "a range (e.g. 1-3), or 'all':"
    )
    while True:
        choice = input("> ").strip().lower()
        if not choice:
            print("Please enter at least one selection.")
            continue
        try:
            selected = _parse_selection(choice, len(files))
        except ValueError as exc:
            print(f"Invalid selection: {exc}")
            continue
        if selected:
            return [files[i] for i in selected]
        print("No files selected. Try again.")


def _parse_selection(choice: str, count: int) -> list[int]:
    if choice == "all":
        return list(range(count))

    indices: list[int] = []
    for part in choice.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = _parse_index(start_s, count)
            end = _parse_index(end_s, count)
            if start > end:
                start, end = end, start
            indices.extend(range(start, end + 1))
        else:
            indices.append(_parse_index(part, count))

    return sorted(set(indices))


def _parse_index(token: str, count: int) -> int:
    idx = int(token) - 1
    if idx < 0 or idx >= count:
        raise ValueError(f"number {token} is out of range (1-{count})")
    return idx


def output_xlsx_path(selected_inputs: list[Path]) -> Path:
    if len(selected_inputs) == 1:
        name = f"{selected_inputs[0].stem}.xlsx"
    else:
        name = "combined_rate_card.xlsx"
    return HARDCODED_OUTPUT_FOLDER / name


def run_pipeline(selected_inputs: list[Path]) -> Path:
    HARDCODED_PROCESSING_FOLDER.mkdir(parents=True, exist_ok=True)
    HARDCODED_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    cleaned_paths: list[Path] = []
    for input_path in selected_inputs:
        print(f"\nCleaning: {input_path.name}")
        cleaned = clean_json_file(input_path, HARDCODED_PROCESSING_FOLDER)
        print(f"  -> {cleaned}")
        cleaned_paths.append(cleaned)

    out_path = output_xlsx_path(selected_inputs)
    print(f"\nBuilding Excel ({len(cleaned_paths)} contract(s))...")
    result = transform_cleaned_jsons_to_xlsx(cleaned_paths, out_path)
    print(f"  -> {result}")
    return result


def main() -> int:
    try:
        files = list_input_json_files()
        selected = prompt_file_selection(files)
        print("\nSelected:")
        for path in selected:
            print(f"  - {path.name}")
        out = run_pipeline(selected)
        print(f"\nDone. Output written to:\n  {out}")
        return 0
    except (KeyboardInterrupt, EOFError):
        print("\nCancelled.")
        return 130
    except Exception as exc:
        print(f"\nError: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
