from __future__ import annotations

import json
import posixpath
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET

MAIN_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_URI = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MAIN_NS = {"a": MAIN_URI}
REL_NS = {"rel": REL_URI}
REL_ID_ATTR = f"{{{OFFICE_REL_URI}}}id"
TEXT_TAG = f"{{{MAIN_URI}}}t"
META_PATTERN = re.compile(r"^\s*convert\(\s*([^,]+?)\s*,\s*([^,]+?)\s*,\s*([^)]+?)\s*\)\s*$")
FIELD_PATTERN = re.compile(r"^\s*(repeated\s+)?([A-Za-z_]\w*)\s+([A-Za-z_]\w*)\s*=\s*\d+\s*,?\s*$")

ET.register_namespace("", MAIN_URI)
ET.register_namespace("r", OFFICE_REL_URI)


@dataclass(frozen=True)
class SheetMeta:
    config_file: str
    output_file: str
    scheme_name: str


@dataclass(frozen=True)
class FieldDef:
    name: str
    repeated: bool


def main() -> int:
    script_path = Path(__file__).resolve()
    repo_root = script_path.parents[2]
    excel_dir = repo_root / "common" / "excel" / "xls"
    schema_dir = repo_root / "common" / "cfg"
    cfg_dir = repo_root / "Assets" / "Cfg"

    schema_cache = {path.name: parse_schema_file(path) for path in schema_dir.glob("*.txt")}
    modified_sheets: list[tuple[str, str, list[str]]] = []

    for workbook_path in sorted(excel_dir.glob("*.xlsx")):
        if should_skip_workbook(workbook_path):
            continue

        sheet_updates: dict[str, list[list[str]]] = {}
        with zipfile.ZipFile(workbook_path) as zf:
            shared_strings = read_shared_strings(zf)
            for sheet_name, sheet_path in read_sheet_refs(zf):
                rows = read_sheet_rows(zf, sheet_path, shared_strings)
                if len(rows) < 2:
                    continue

                meta = parse_sheet_meta(rows[0][0] if rows[0] else "")
                if meta is None:
                    continue

                fields = schema_cache.get(meta.config_file, {}).get(meta.scheme_name)
                if not fields:
                    continue

                headers = [value.strip() for value in rows[1]]
                missing_headers = [field.name for field in fields if field.name not in headers]
                if not missing_headers:
                    continue

                existing_records = load_existing_records(cfg_dir / meta.output_file, meta.scheme_name, fields[0].name)
                updated_rows = build_updated_rows(rows, fields, existing_records)
                sheet_updates[sheet_path] = updated_rows
                modified_sheets.append((workbook_path.name, sheet_name, missing_headers))

        if sheet_updates:
            patch_workbook(workbook_path, sheet_updates)

    if not modified_sheets:
        print("[ExcelHeaderSync] 未检测到缺失表头。")
        return 0

    print("[ExcelHeaderSync] 已完成以下工作表的表头补全：")
    for workbook_name, sheet_name, missing_headers in modified_sheets:
        print(f"  - {workbook_name}::{sheet_name} -> {', '.join(missing_headers)}")

    return 0


def should_skip_workbook(path: Path) -> bool:
    name = path.name.lower()
    return name.startswith("~$") or name == "equipment_new.xlsx"


def parse_schema_file(schema_path: Path) -> dict[str, list[FieldDef]]:
    schemes: dict[str, list[FieldDef]] = {}
    current_scheme: str | None = None
    current_fields: list[FieldDef] = []

    for raw_line in schema_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.split("//", 1)[0].strip()
        if not line:
            continue

        if current_scheme is None:
            if line.endswith("{"):
                current_scheme = line[:-1].strip()
                current_fields = []
            continue

        if line == "}":
            schemes[current_scheme] = current_fields
            current_scheme = None
            current_fields = []
            continue

        match = FIELD_PATTERN.match(line)
        if not match:
            continue

        current_fields.append(FieldDef(name=match.group(3), repeated=bool(match.group(1))))

    return schemes


def parse_sheet_meta(value: str) -> SheetMeta | None:
    match = META_PATTERN.match(value.strip())
    if not match:
        return None

    return SheetMeta(
        config_file=match.group(1).strip(),
        output_file=match.group(2).strip(),
        scheme_name=match.group(3).strip(),
    )


def load_existing_records(output_path: Path, scheme_name: str, key_field: str) -> dict[str, dict[str, object]]:
    if not output_path.exists():
        return {}

    try:
        payload = json.loads(output_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}

    rows = payload.get(scheme_name)
    if not isinstance(rows, list):
        return {}

    records: dict[str, dict[str, object]] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        key = row.get(key_field)
        if key is None:
            continue
        records[str(key)] = row
    return records


def build_updated_rows(
    rows: list[list[str]],
    fields: list[FieldDef],
    existing_records: dict[str, dict[str, object]],
) -> list[list[str]]:
    current_headers = [value.strip() for value in rows[1]]
    key_field = fields[0].name
    updated_rows: list[list[str]] = [list(rows[0]), [field.name for field in fields]]

    for row in rows[2:]:
        if is_blank_row(row):
            updated_rows.append([])
            continue

        row_map: dict[str, str] = {}
        for index, header in enumerate(current_headers):
            if not header:
                continue
            row_map[header] = row[index] if index < len(row) else ""

        record_key = row_map.get(key_field, "").strip()
        fallback_record = existing_records.get(record_key, {})

        normalized_row: list[str] = []
        for field in fields:
            value = row_map.get(field.name, "")
            if value == "" and field.name in fallback_record:
                value = normalize_record_value(fallback_record[field.name], field.repeated)
            normalized_row.append(value)

        while normalized_row and normalized_row[-1] == "":
            normalized_row.pop()
        updated_rows.append(normalized_row)

    return updated_rows


def normalize_record_value(value: object, repeated: bool) -> str:
    if value is None:
        return ""

    if repeated:
        if isinstance(value, list):
            return ",".join(str(item) for item in value)
        return str(value)

    return str(value)


def is_blank_row(row: list[str]) -> bool:
    return all(cell.strip() == "" for cell in row)


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.iter(TEXT_TAG)) for item in root.findall("a:si", MAIN_NS)]


def read_sheet_refs(zf: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: normalize_zip_path(rel.attrib["Target"])
        for rel in rels_root.findall("rel:Relationship", REL_NS)
    }

    refs: list[tuple[str, str]] = []
    sheets_root = workbook_root.find("a:sheets", MAIN_NS)
    if sheets_root is None:
        return refs

    for sheet in sheets_root.findall("a:sheet", MAIN_NS):
        rel_id = sheet.attrib.get(REL_ID_ATTR)
        if rel_id and rel_id in rel_map:
            refs.append((sheet.attrib.get("name", ""), rel_map[rel_id]))

    return refs


def normalize_zip_path(target: str) -> str:
    normalized = target.replace("\\", "/").lstrip("/")
    if not normalized.startswith("xl/"):
        normalized = f"xl/{normalized}"
    return posixpath.normpath(normalized)


def read_sheet_rows(zf: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(zf.read(sheet_path))
    sheet_data = root.find("a:sheetData", MAIN_NS)
    if sheet_data is None:
        return []

    rows: list[list[str]] = []
    for row_node in sheet_data.findall("a:row", MAIN_NS):
        row_values: dict[int, str] = {}
        max_index = -1

        for cell in row_node.findall("a:c", MAIN_NS):
            ref = cell.attrib.get("r", "")
            column_index = column_index_from_ref(ref) if ref else max_index + 1
            row_values[column_index] = read_cell_value(cell, shared_strings)
            max_index = max(max_index, column_index)

        if max_index < 0:
            rows.append([])
            continue

        row = [""] * (max_index + 1)
        for index, value in row_values.items():
            row[index] = value

        while row and row[-1] == "":
            row.pop()
        rows.append(row)

    return rows


def read_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("a:v", MAIN_NS)

    if cell_type == "s" and value_node is not None:
        index = int(value_node.text or "0")
        return shared_strings[index] if 0 <= index < len(shared_strings) else ""

    if cell_type == "inlineStr":
        inline_node = cell.find("a:is", MAIN_NS)
        if inline_node is None:
            return ""
        return "".join(node.text or "" for node in inline_node.iter(TEXT_TAG))

    if value_node is not None and value_node.text is not None:
        return value_node.text

    return ""


def column_index_from_ref(cell_ref: str) -> int:
    column = 0
    for char in cell_ref:
        if not char.isalpha():
            break
        column = column * 26 + (ord(char.upper()) - ord("A") + 1)
    return max(column - 1, 0)


def patch_workbook(workbook_path: Path, updates: dict[str, list[list[str]]]) -> None:
    with zipfile.ZipFile(workbook_path) as source_zip:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            temp_path = Path(temp_file.name)

        try:
            with zipfile.ZipFile(temp_path, "w") as target_zip:
                for info in source_zip.infolist():
                    data = source_zip.read(info.filename)
                    if info.filename in updates:
                        data = patch_sheet_xml(data, updates[info.filename])
                    target_zip.writestr(info, data)

            workbook_path.write_bytes(temp_path.read_bytes())
        finally:
            if temp_path.exists():
                temp_path.unlink()


def patch_sheet_xml(sheet_bytes: bytes, rows: list[list[str]]) -> bytes:
    root = ET.fromstring(sheet_bytes)
    sheet_data = root.find("a:sheetData", MAIN_NS)
    if sheet_data is None:
        return sheet_bytes

    existing_rows = list(sheet_data.findall("a:row", MAIN_NS))
    existing_map = {
        int(row.attrib.get("r", str(index + 1))): row
        for index, row in enumerate(existing_rows)
    }
    template_styles = {
        row_number: [cell.attrib.get("s") for cell in row.findall("a:c", MAIN_NS)]
        for row_number, row in existing_map.items()
    }

    for row in existing_rows:
        sheet_data.remove(row)

    max_columns = max((len(row) for row in rows), default=1)

    for row_number, row_values in enumerate(rows, start=1):
        row_node = existing_map.get(row_number, ET.Element(f"{{{MAIN_URI}}}row"))
        for child in list(row_node):
            row_node.remove(child)

        row_node.set("r", str(row_number))
        if row_values:
            row_node.set("spans", f"1:{len(row_values)}")
        elif "spans" in row_node.attrib:
            del row_node.attrib["spans"]

        style_row = template_styles.get(row_number, [])
        fallback_style = next((style for style in reversed(style_row) if style), None)
        for column_number, value in enumerate(row_values, start=1):
            style_value = style_row[column_number - 1] if column_number - 1 < len(style_row) else fallback_style
            row_node.append(build_inline_cell(column_number, row_number, value, style_value))

        sheet_data.append(row_node)

    dimension = root.find("a:dimension", MAIN_NS)
    if dimension is not None:
        last_ref = f"{column_name_from_index(max_columns)}{max(len(rows), 1)}"
        dimension.set("ref", f"A1:{last_ref}")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_inline_cell(column_number: int, row_number: int, value: str, style_value: str | None) -> ET.Element:
    cell = ET.Element(f"{{{MAIN_URI}}}c")
    cell.set("r", f"{column_name_from_index(column_number)}{row_number}")
    cell.set("t", "inlineStr")
    if style_value:
        cell.set("s", style_value)

    inline_node = ET.SubElement(cell, f"{{{MAIN_URI}}}is")
    text_node = ET.SubElement(inline_node, TEXT_TAG)
    if value.startswith(" ") or value.endswith(" ") or "\n" in value:
        text_node.set(XML_SPACE, "preserve")
    text_node.text = value
    return cell


def column_name_from_index(index: int) -> str:
    if index <= 0:
        return "A"

    chars: list[str] = []
    current = index
    while current > 0:
        current -= 1
        chars.append(chr(ord("A") + (current % 26)))
        current //= 26
    return "".join(reversed(chars))


if __name__ == "__main__":
    raise SystemExit(main())
