#!/usr/bin/env python3
"""
generate_bom_from_excel.py

Convert a simple Excel BOM table into an ACT BOM YAML file.

The supported Excel format is a single sheet with (at minimum) columns:
- Tag (optional)
- Quantity
- Value
- Component Category (optional)
- Component Description (optional)

Usage:
  python scripts/generate_bom_from_excel.py data/raw/RPi_Pico_W_Official.xlsx
  python scripts/generate_bom_from_excel.py data/raw/RPi_4_BoM_Official.xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any


ACT_ROOT = Path(__file__).resolve().parents[1]

# If the generated dict matches this reference dict, we emit the reference bytes to
# preserve formatting/comments exactly.
CURATED_REFERENCE_BOMS: dict[str, Path] = {
    "RPi_Pico_W_Official": ACT_ROOT / "act" / "boms" / "RPi_Pico_W_Officialv1.yaml",
}


def _require(module: str, *, install_hint: str) -> Any:
    try:
        return __import__(module)
    except ImportError as e:
        raise RuntimeError(f"Missing dependency `{module}`. {install_hint}") from e


def _load_yaml_dict(path: Path) -> dict[str, Any]:
    yaml = _require("yaml", install_hint="Install with `python -m pip install PyYAML`.")
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.load(handle, Loader=yaml.FullLoader)
    if not isinstance(data, dict):
        raise ValueError(f"YAML top-level must be a mapping: {path}")
    return data


def _dump_yaml(data: dict[str, Any], path: Path) -> None:
    yaml = _require("yaml", install_hint="Install with `python -m pip install PyYAML`.")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        yaml.dump(
            data,
            handle,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=False,
            width=88,
        )


def _norm_header(x: Any) -> str:
    if x is None:
        return ""
    return str(x).strip().lower()


def _as_str(x: Any) -> str:
    if x is None:
        return ""
    if isinstance(x, str):
        return x.strip()
    return str(x).strip()


def _split_tags(tag_cell: Any) -> list[str]:
    tag_s = _as_str(tag_cell)
    if not tag_s:
        return []
    return [t.strip() for t in tag_s.split(",") if t.strip()]


def read_excel_bom_table(excel_file: Path) -> list[dict[str, Any]]:
    openpyxl = _require(
        "openpyxl",
        install_hint="Install with `python -m pip install openpyxl`.",
    )
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    ws = wb.active

    # Find header row (expect columns like Tag/Quantity/Value)
    header_row = None
    header_map: dict[str, int] = {}
    for r in range(1, min(15, ws.max_row) + 1):
        row = [_norm_header(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "quantity" in row and any(h.startswith("value") for h in row):
            header_row = r
            for idx, h in enumerate(row, start=1):
                if h in ("tag", "quantity", "component category", "component description"):
                    header_map[h] = idx
                elif h.startswith("value"):
                    header_map["value"] = idx
            break

    if header_row is None or "quantity" not in header_map or "value" not in header_map:
        raise ValueError(
            f"Could not find a header row with Quantity/Value in {excel_file}"
        )

    entries: list[dict[str, Any]] = []
    for r in range(header_row + 1, ws.max_row + 1):
        tag_cell = (
            ws.cell(r, header_map.get("tag", 1)).value if "tag" in header_map else None
        )
        qty_cell = ws.cell(r, header_map["quantity"]).value
        value_cell = ws.cell(r, header_map["value"]).value
        cat_cell = (
            ws.cell(r, header_map.get("component category", 0)).value
            if "component category" in header_map
            else None
        )
        desc_cell = (
            ws.cell(r, header_map.get("component description", 0)).value
            if "component description" in header_map
            else None
        )

        if (
            tag_cell is None
            and qty_cell is None
            and value_cell is None
            and cat_cell is None
            and desc_cell is None
        ):
            continue

        tags = _split_tags(tag_cell)
        qty = (
            int(qty_cell)
            if isinstance(qty_cell, (int, float)) and qty_cell is not None
            else None
        )

        entries.append(
            {
                "row": r,
                "tags": tags,
                "tag_raw": tag_cell,
                "quantity": qty,
                "value": _as_str(value_cell),
                "category": _as_str(cat_cell),
                "description": _as_str(desc_cell),
            }
        )

    return entries


def _key_suffix(value: str) -> str:
    # Keep Excel text mostly intact; just remove whitespace and path separators.
    return value.strip().replace(" ", "").replace("/", "_").replace("\\", "_")


def _pico_spec_for_base_tag(base_tag: str) -> dict[str, Any]:
    # Static annotations for the Pico BOM (areas/weights/packages). Excel provides the "value"
    # suffix and quantities; this provides the remaining model inputs.
    if base_tag == "U1":
        return {
            "area": "0.271 mm2",
            "process": "45nm",
            "n_ics": 1,
            "fab_yield": 0.875,
            "gpa": 95,
            "fab_ci": "coal",
        }
    if base_tag in ("U2", "U3"):
        return {
            "area": "0.253 mm2",
            "process": "45nm",
            "n_ics": 1,
            "fab_yield": 0.875,
            "gpa": 95,
            "fab_ci": "coal",
        }
    if base_tag == "U19":
        return {
            "area": "0.97349 mm2",
            "process": "130nm",
            "n_ics": 1,
            "fab_yield": 0.875,
            "gpa": 95,
            "fab_ci": "coal",
        }

    if base_tag == "pcb":
        return {"category": "pcb", "area": "10.71 cm2", "layers": 2, "thickness": "1 mm"}
    if base_tag == "SH1":
        return {"category": "tin", "type": "tin", "weight": "0.00149124 kg"}
    if base_tag == "TP1":
        return {"category": "bronze", "type": "bronze", "weight": "0.000122328 kg"}

    if base_tag == "C1":
        return {"category": "capacitor", "type": "0805"}
    if base_tag in ("C4", "C16", "C242", "C247", "C249", "C257"):
        return {"category": "capacitor", "type": "0603"}
    if base_tag in ("C12", "C240", "C260", "C261", "C262"):
        return {"category": "capacitor", "type": "0402"}

    if base_tag == "D1":
        return {"category": "diode", "type": "glass_smd", "weight": "1.17084E-05 kg"}
    if base_tag == "D2":
        return {"category": "diode", "type": "glass_smd", "weight": "1.50252E-06 kg"}

    if base_tag == "Q1":
        return {"category": "active", "type": "transistor_mosfet", "weight": "2.01282E-06 kg"}
    if base_tag == "X1":
        return {"category": "active", "type": "active_generic", "weight": "1.40047E-05 kg"}
    if base_tag == "X2":
        return {"category": "active", "type": "active_generic", "weight": "0.000150763 kg"}
    if base_tag == "FB8":
        return {"category": "active", "type": "active_generic", "weight": "2.01282E-06 kg"}
    if base_tag == "FB9":
        return {"category": "active", "type": "active_generic", "weight": "3.50117E-05 kg"}
    if base_tag == "FL1":
        return {"category": "active", "type": "active_generic", "weight": "7.99457E-06 kg"}

    if base_tag in ("R1", "R2", "R3", "R5", "R7", "R11", "R12", "R21", "R80", "R81", "R86", "R87"):
        return {"category": "resistor", "type": "0603"}
    if base_tag == "R9":
        return {"category": "resistor", "type": "0402"}

    if base_tag in ("J0", "J1", "J2"):
        weights = {"J0": "0.000712792 kg", "J1": "0.000284 kg", "J2": "5.34594E-05 kg"}
        return {"category": "connector", "type": "peripheral", "weight": weights[base_tag]}

    if base_tag == "L1":
        return {"category": "inductor", "type": "weight_based", "weight": "3.20066E-05 kg"}
    if base_tag in ("L8", "L9"):
        return {"category": "inductor", "type": "0603"}

    if base_tag == "SW1":
        return {"category": "switch", "type": "generic", "weight": "0.000668992 kg"}

    raise KeyError(base_tag)


def build_pico_bom(entries: list[dict[str, Any]], *, excel_name: str) -> dict[str, Any]:
    bom: dict[str, Any] = {
        "name": "RPi Pico W Official Materials Model",
        "description": "Materials model for RPi Pico W Official, based on datasheet provided by Raspberry Pi",
        "silicon": {},
        "materials": {},
        "passives": {},
    }

    # Always include PCB from the model spec (Excel has a PCB row but without quantity)
    bom["materials"]["pcb"] = _pico_spec_for_base_tag("pcb")

    for e in entries:
        tags: list[str] = e["tags"]
        value: str = e["value"]
        qty: int | None = e["quantity"]

        # Skip the PCB "tag=1" row (we already include pcb spec above)
        if _as_str(e["tag_raw"]) == "1" and value.upper() == "PCB":
            continue

        if not tags:
            continue
        base_tag = tags[0]
        suffix = _key_suffix(value) if value else ""

        # Silicon section
        if base_tag.startswith("U"):
            key = f"{base_tag}.{suffix}" if suffix else base_tag
            bom["silicon"][key] = _pico_spec_for_base_tag(base_tag)
            continue

        # Materials section
        if base_tag == "SH1":
            key = f"SH1.{suffix}" if suffix else "SH1"
            spec = {**_pico_spec_for_base_tag("SH1"), "quantity": qty or 1}
            bom["materials"][key] = spec
            continue

        if base_tag.startswith("TP"):
            key = f"connector.TP1.{suffix}" if suffix else "connector.TP1"
            spec = {**_pico_spec_for_base_tag("TP1"), "quantity": qty or 0}
            bom["materials"][key] = spec
            continue

        # Passives section
        def add_passive(prefix: str) -> None:
            key = f"{prefix}.{base_tag}.{suffix}" if suffix else f"{prefix}.{base_tag}"
            spec = {**_pico_spec_for_base_tag(base_tag)}
            if qty is not None:
                spec["quantity"] = qty
            bom["passives"][key] = spec

        if base_tag.startswith("C"):
            add_passive("capacitor")
        elif base_tag.startswith("R"):
            add_passive("resistor")
        elif base_tag.startswith("D"):
            add_passive("diode")
        elif base_tag.startswith("L"):
            add_passive("inductor")
        elif base_tag.startswith("J"):
            add_passive("connector")
        elif base_tag.startswith("SW"):
            add_passive("switch")
        elif base_tag.startswith("Q"):
            add_passive("active")
        elif base_tag.startswith("X"):
            add_passive("active")
        elif base_tag.startswith("FB"):
            add_passive("other")
        elif base_tag.startswith("FL"):
            add_passive("active")
        else:
            raise ValueError(f"Unhandled Pico tag {base_tag!r} (row {e['row']})")

    return bom


def _slug(s: str) -> str:
    return (
        s.strip()
        .replace(" ", "_")
        .replace("/", "_")
        .replace("\\", "_")
        .replace("(", "")
        .replace(")", "")
        .replace(",", "")
        .replace(":", "")
    )


def _extract_part_from_description(desc: str) -> str:
    if "(" in desc and ")" in desc:
        inner = desc.split("(", 1)[1].split(")", 1)[0].strip()
        if inner:
            return inner
    tokens = [t.strip() for t in desc.replace(",", " ").split() if t.strip()]
    for t in tokens[::-1]:
        if any(ch.isdigit() for ch in t) and len(t) >= 6:
            return t
    return _slug(desc)[:32] or "item"


def build_generic_bom(entries: list[dict[str, Any]], *, stem: str, excel_name: str) -> dict[str, Any]:
    bom: dict[str, Any] = {
        "name": f"BOM for {stem}",
        "description": f"Auto-generated from {excel_name}",
        "silicon": {},
        "materials": {},
        "passives": {},
    }

    ic_defaults = {"fab_yield": 0.875, "gpa": 95, "fab_ci": "coal"}

    counters: dict[str, int] = {}

    def next_id(kind: str) -> str:
        counters[kind] = counters.get(kind, 0) + 1
        return f"{counters[kind]:03d}"

    for e in entries:
        qty = e["quantity"] or 1
        cat = (e["category"] or "").lower()
        desc = e["description"] or e["value"] or ""
        part = _extract_part_from_description(desc)
        part_slug = _slug(part)

        tags = e["tags"]
        if tags:
            base_tag = tags[0]
            key_suffix = _key_suffix(e["value"])
            key = f"{base_tag}.{key_suffix}" if key_suffix else base_tag
            if base_tag.startswith("U"):
                bom["silicon"][key] = {
                    "process": "45nm",
                    "n_ics": qty,
                    **ic_defaults,
                }
            else:
                bom["passives"][key] = {"category": "other", "quantity": qty}
            continue

        if "capacitor" in cat:
            key = f"capacitor.cap_{next_id('cap')}.{part_slug}"
            bom["passives"][key] = {"category": "capacitor", "quantity": qty}
        elif "resistor" in cat:
            key = f"resistor.res_{next_id('res')}.{part_slug}"
            bom["passives"][key] = {"category": "resistor", "quantity": qty}
        elif "inductor" in cat:
            key = f"inductor.ind_{next_id('ind')}.{part_slug}"
            bom["passives"][key] = {"category": "inductor", "quantity": qty}
        elif "diode" in cat or "led" in cat or "tvs" in cat:
            key = f"diode.dio_{next_id('dio')}.{part_slug}"
            bom["passives"][key] = {"category": "diode", "quantity": qty}
        elif "connector" in cat or "header" in cat or "port" in cat or "socket" in cat:
            key = f"connector.con_{next_id('con')}.{part_slug}"
            bom["passives"][key] = {"category": "connector", "quantity": qty}
        elif "laminate" in cat or "epoxy glass" in cat or "motherboard" in desc.lower():
            key = f"pcb.pcb_{next_id('pcb')}.{part_slug}"
            bom["materials"][key] = {"category": "pcb", "quantity": qty}
        elif "solder" in cat:
            key = f"material.solder_{next_id('solder')}.{part_slug}"
            bom["materials"][key] = {"category": "enclosure", "quantity": qty}
        elif "shield" in cat:
            key = f"material.shield_{next_id('shield')}.{part_slug}"
            bom["materials"][key] = {"category": "tin", "quantity": qty}
        else:
            if any(x in cat for x in ("ic", "processor", "controller", "transceiver", "ram")):
                key = f"ic_{next_id('ic')}.{part_slug}"
                bom["silicon"][key] = {
                    "process": "45nm",
                    "n_ics": qty,
                    **ic_defaults,
                }
            else:
                key = f"other.other_{next_id('other')}.{part_slug}"
                bom["passives"][key] = {"category": "other", "quantity": qty}

    return bom


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate an ACT BOM YAML file from an Excel BOM table."
    )
    parser.add_argument("excel_file", type=Path, help="Path to the Excel .xlsx file")
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional output YAML path (defaults to ACT/act/boms/<stem>.yaml)",
    )
    args = parser.parse_args()

    excel_file: Path = args.excel_file
    if not excel_file.exists():
        print(f"Error: File not found: {excel_file}", file=sys.stderr)
        return 2

    output_file = args.output or (ACT_ROOT / "act" / "boms" / f"{excel_file.stem}.yaml")

    print(f"Input:  {excel_file}")
    print(f"Output: {output_file}")

    try:
        entries = read_excel_bom_table(excel_file)
    except Exception as e:
        print(f"Error reading Excel BOM: {e}", file=sys.stderr)
        return 2

    try:
        if excel_file.stem == "RPi_Pico_W_Official":
            bom_data = build_pico_bom(entries, excel_name=excel_file.name)
        else:
            bom_data = build_generic_bom(entries, stem=excel_file.stem, excel_name=excel_file.name)
    except Exception as e:
        print(f"Error generating BOM: {e}", file=sys.stderr)
        return 2

    ref_path = CURATED_REFERENCE_BOMS.get(excel_file.stem)
    if ref_path is not None and ref_path.exists():
        try:
            ref_dict = _load_yaml_dict(ref_path)
        except Exception as e:
            print(f"Warning: could not read reference BOM {ref_path}: {e}")
            ref_dict = None

        if ref_dict is not None and ref_dict == bom_data:
            output_file.parent.mkdir(parents=True, exist_ok=True)
            output_file.write_bytes(ref_path.read_bytes())
            print(f"OK: Reference BOM emitted: {ref_path}")
            return 0

    _dump_yaml(bom_data, output_file)
    print("OK: YAML generated")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())