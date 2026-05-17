"""Extract Document Intelligence fields from input JSON into cleaned output."""

from __future__ import annotations

import json
import sys
from pathlib import Path

INPUT_DIR = Path(__file__).resolve().parent / "input"
OUTPUT_DIR = Path(__file__).resolve().parent / "processing"

VALUE_KEYS = (
    "valueString",
    "valueNumber",
    "valueInteger",
    "valueDate",
    "valueTime",
    "valuePhoneNumber",
    "valueCountryRegion",
    "valueSelectionMark",
    "valueSignature",
    "valueCurrency",
    "valueAddress",
    "content",
)


def _has_scalar_value(field: dict) -> bool:
    return any(field.get(key) is not None for key in VALUE_KEYS)


def extract_field_value(field: dict | None) -> object | None:
    """Convert an Azure Document Intelligence field node to a plain value."""
    if not field or not isinstance(field, dict):
        return None

    field_type = field.get("type")

    if field_type == "array":
        items = field.get("valueArray") or []
        cleaned = [extract_field_value(item) for item in items]
        return [item for item in cleaned if item is not None]

    if field_type == "object":
        obj = field.get("valueObject") or {}
        cleaned: dict[str, object] = {}
        for key, child in obj.items():
            value = extract_field_value(child)
            if value is not None:
                cleaned[key] = value
        return cleaned or None

    for key in VALUE_KEYS:
        if field.get(key) is not None:
            return field[key]

    return None


def extract_fields_from_document(doc: dict) -> dict[str, object]:
    """Return only field names that have extractable values."""
    raw_fields = doc.get("fields") or {}
    cleaned: dict[str, object] = {}
    for name, field in raw_fields.items():
        value = extract_field_value(field)
        if value is not None:
            if isinstance(value, list) and not value:
                continue
            if isinstance(value, dict) and not value:
                continue
            cleaned[name] = value
    return cleaned


def clean_analyze_result(data: dict, source_name: str) -> dict:
    analyze = data.get("analyzeResult") or {}
    documents = analyze.get("documents") or []

    if not documents:
        raise ValueError(f"No documents found in analyzeResult for {source_name}")

    output_documents = []
    for doc in documents:
        fields = extract_fields_from_document(doc)
        if not fields:
            continue
        entry: dict[str, object] = {"fields": fields}
        if doc.get("docType"):
            entry["docType"] = doc["docType"]
        output_documents.append(entry)

    if not output_documents:
        raise ValueError(f"No extractable fields found for {source_name}")

    result: dict[str, object] = {
        "source": source_name,
        "status": data.get("status"),
        "documents": output_documents,
    }
    if len(output_documents) == 1:
        result["fields"] = output_documents[0]["fields"]
        if output_documents[0].get("docType"):
            result["docType"] = output_documents[0]["docType"]
    return result


def process_file(input_path: Path, output_dir: Path) -> Path:
    with input_path.open(encoding="utf-8") as f:
        data = json.load(f)

    cleaned = clean_analyze_result(data, input_path.name)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / input_path.name
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(cleaned, f, indent=2, ensure_ascii=False)
        f.write("\n")
    return output_path


def main(argv: list[str] | None = None) -> int:
    args = argv if argv is not None else sys.argv[1:]
    input_dir = Path(args[0]) if args else INPUT_DIR
    output_dir = Path(args[1]) if len(args) > 1 else OUTPUT_DIR

    json_files = sorted(input_dir.glob("*.json"))
    if not json_files:
        print(f"No JSON files found in {input_dir}", file=sys.stderr)
        return 1

    for path in json_files:
        out = process_file(path, output_dir)
        print(f"Wrote {out}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
