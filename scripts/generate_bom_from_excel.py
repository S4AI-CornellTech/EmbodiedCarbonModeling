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
import math
import re
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


def _is_missing(v: Any) -> bool:
    """Return True if value is None, empty string, or NaN float."""
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def _to_float(v: Any) -> float | None:
    """Convert a value to float, returning None if not possible or missing."""
    if _is_missing(v):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _convert_process_node(value: Any) -> str:
    """Convert 28.0 or '28' to '28nm'; pass through 'lpddr2_20nm' style strings."""
    if _is_missing(value):
        return "45nm"
    value_str = str(value).strip()
    if "nm" in value_str or "_" in value_str:
        return value_str
    try:
        num = int(float(value_str))
        return f"{num}nm"
    except (TypeError, ValueError):
        return "45nm"


def _area_str(area_mm2: float) -> str | None:
    """Return area as 'X mm2' for small ICs or 'X.XXXX cm2' for large dies/PCBs."""
    if area_mm2 is None or (isinstance(area_mm2, float) and math.isnan(area_mm2)):
        return None
    if area_mm2 <= 0:
        return None
    if area_mm2 < 100:
        return f"{area_mm2} mm2"
    return f"{area_mm2 / 100:.4f} cm2"


_STANDARD_PACKAGES = {"0201", "0402", "0603", "0805"}


def _fmt_weight(kg: float) -> str:
    """Format a weight value cleanly, e.g. 0.0016899999999999999 → '0.00169 kg'."""
    return f"{kg:.6g} kg"


def _extract_package_size(pkg_raw: Any, desc: str) -> str:
    """Extract standard package size code from package column or description string."""
    if not _is_missing(pkg_raw):
        pkg_clean = str(pkg_raw).strip().replace("mm", "").replace(" ", "")
        if pkg_clean in _STANDARD_PACKAGES:
            return pkg_clean

    search_str = desc or ""
    for pkg in _STANDARD_PACKAGES:
        if re.search(r"\b" + pkg + r"\b", search_str):
            return pkg
    return "0603"


def read_excel_bom_table(excel_file: Path) -> list[dict[str, Any]]:
    openpyxl = _require(
        "openpyxl",
        install_hint="Install with `python -m pip install openpyxl`.",
    )
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    ws = wb.active

    # Normalised header text → internal field name.
    # Duplicate column headers (e.g. two "SoC technology nodes" columns) are
    # handled below by appending "_alt" for the second occurrence.
    _HEADER_TO_FIELD: dict[str, str] = {
        "tag": "tag",
        "quantity": "quantity",
        "component category": "category",           # backward compat (Pico sheet)
        "component category(act)": "category_act",  # ACT-specific category (preferred)
        "component category(manufacturer)": "category_mfr",
        "component description": "description",
        "weight (kg)": "weight_kg",
        "part number": "part_number",
        "area (sq mm)": "area_sq_mm",
        "layers": "layers",
        "thickness (mm)": "thickness_mm",
        "package size (mm)": "package_size_mm",
        "package to die ratio": "package_to_die_ratio",
        "die size (mm2)": "die_size_mm2",
        "soc technology nodes": "soc_process",
        "gb": "gb",
        "capacity": "capacity",
        "process": "process",
        "memory emissions factor g co2e": "memory_emission_factor",
        "type": "type",
        "packaging": "packaging",
    }

    # Find header row; col_map maps field_name → column index (1-based)
    header_row = None
    col_map: dict[str, int] = {}
    for r in range(1, min(15, ws.max_row) + 1):
        row = [_norm_header(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "quantity" not in row or not any(h.startswith("value") for h in row):
            continue

        header_row = r
        seen_raw: dict[str, int] = {}
        for idx, h in enumerate(row, start=1):
            if not h:
                continue
            seen_raw[h] = seen_raw.get(h, 0) + 1
            occurrence = seen_raw[h]

            if h.startswith("value"):
                col_map.setdefault("value", idx)
                continue

            if h in _HEADER_TO_FIELD:
                field = _HEADER_TO_FIELD[h]
                if occurrence == 1:
                    col_map.setdefault(field, idx)
                else:
                    col_map.setdefault(field + "_alt", idx)
        break

    if header_row is None or "quantity" not in col_map or "value" not in col_map:
        raise ValueError(
            f"Could not find a header row with Quantity/Value in {excel_file}"
        )

    def _cell(r: int, field: str) -> Any:
        idx = col_map.get(field)
        return ws.cell(r, idx).value if idx else None

    entries: list[dict[str, Any]] = []
    for r in range(header_row + 1, ws.max_row + 1):
        tag_cell = _cell(r, "tag")
        qty_cell = _cell(r, "quantity")
        value_cell = _cell(r, "value")
        cat_act_cell = _cell(r, "category_act")   # Component Category(ACT)
        cat_cell = _cell(r, "category")            # backward-compat generic header
        desc_cell = _cell(r, "description")

        # Skip completely empty rows
        if all(_is_missing(v) for v in [tag_cell, qty_cell, value_cell, cat_act_cell, cat_cell, desc_cell]):
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
                # Prefer the ACT-specific category column when available
                "category": _as_str(cat_act_cell or cat_cell),
                "category_mfr": _as_str(_cell(r, "category_mfr")),
                "description": _as_str(desc_cell),
                # Extended fields
                "weight_kg": _to_float(_cell(r, "weight_kg")),
                "part_number": _as_str(_cell(r, "part_number")),
                "area_sq_mm": _to_float(_cell(r, "area_sq_mm")),
                "layers": _to_float(_cell(r, "layers")),
                "thickness_mm": _to_float(_cell(r, "thickness_mm")),
                "package_size_mm": _cell(r, "package_size_mm"),
                "package_to_die_ratio": _to_float(_cell(r, "package_to_die_ratio")),
                "die_size_mm2": _to_float(_cell(r, "die_size_mm2")),
                "soc_process": _cell(r, "soc_process"),
                "soc_process_alt": _cell(r, "soc_process_alt"),
                "gb": _to_float(_cell(r, "gb")),
                "capacity": _to_float(_cell(r, "capacity")),
                "process": _as_str(_cell(r, "process")),
                "memory_emission_factor": _to_float(_cell(r, "memory_emission_factor")),
                "type": _as_str(_cell(r, "type")),
                "packaging": _as_str(_cell(r, "packaging")),
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

    def _ic_key(prefix: str, tags: list[str], value: str, part_slug: str) -> str:
        """Build a silicon key from tag or slug, with counter fallback."""
        if tags:
            tag = tags[0]
            v_slug = _key_suffix(value) if value else ""
            return f"{prefix}.{tag}.{v_slug}" if v_slug else f"{prefix}.{tag}"
        if part_slug:
            return f"{prefix}.{part_slug}"
        return f"{prefix}.main{next_id(prefix.rsplit('.', 1)[-1])}"

    def _passive_key(prefix: str, counter_key: str, tags: list[str]) -> str:
        """Build a passives key from tag or counter."""
        if tags:
            return f"{prefix}.{tags[0]}"
        return f"{prefix}.{counter_key}_{next_id(counter_key)}"

    for e in entries:
        qty = e["quantity"] or 1
        cat = (e["category"] or "").lower()
        desc = e["description"] or e["value"] or ""
        tags = e["tags"]
        value = e["value"] or ""

        # Extended fields
        die_size_mm2 = e.get("die_size_mm2")
        # Prefer the second SoC technology nodes column (more specific) when present
        soc_process_raw = e.get("soc_process_alt") or e.get("soc_process")
        mem_process = e.get("process") or ""
        area_sq_mm = e.get("area_sq_mm")
        # For IC/CPU die area: Die Size (mm2) is authoritative; Area (sq mm) is the
        # fallback because some BOMs store the die area there instead of a PCB area.
        ic_die_mm2 = die_size_mm2 if die_size_mm2 is not None else area_sq_mm
        layers = e.get("layers")
        thickness_mm = e.get("thickness_mm")
        package_size_mm = e.get("package_size_mm")
        weight_kg = e.get("weight_kg")
        capacity_gb = e.get("gb") or e.get("capacity")

        part_slug = _slug(_extract_part_from_description(desc))

        # ── Classify component ───────────────────────────────────────────────────
        desc_lower = desc.lower()
        cat_desc = cat + " " + desc_lower  # combined for keyword matching
        cat_stripped = cat.strip()          # bare ACT category value (e.g. "ic", "dram")

        # When the ACT category column provides an explicit value, trust it entirely and
        # skip keyword heuristics. This prevents false positives like "ceramic" matching
        # "ram", or "NOR Flash" matching "flash" when ACT explicitly says "IC".
        _has_cat = bool(cat_stripped)

        # Material flags are evaluated first so they can exclude IC mis-classification.
        is_pcb = cat_stripped == "pcb" or (not _has_cat and any(
            x in cat_desc
            for x in ("laminate", "pcb", "motherboard", "epoxy glass", "printed circuit")
        ))
        is_solder = "solder" in cat   # matches "solder", "pb free solder", etc.
        is_shield = cat_stripped == "shield" or (not _has_cat and "shield" in cat and not is_solder)
        is_aluminum = (not is_solder) and (not is_shield) and (
            cat_stripped == "aluminum"
            or (not _has_cat and any(x in cat_desc for x in ("alumin", "aluminum", "aluminium")))
        )
        is_other = cat_stripped == "other"

        # When ACT category is explicitly "ic", treat as silicon IC regardless of keywords.
        _act_ic = cat_stripped == "ic"

        is_ram = cat_stripped == "dram" or (not _has_cat and any(
            x in cat_desc for x in ("ram", "dram", "lpddr", "sdram", "memory")
        ))
        is_flash = (not is_ram) and (cat_stripped == "flash" or (not _has_cat and any(
            x in cat_desc for x in ("flash", "nand", "emmc", "ssd", "storage")
        )))
        is_cpu = (not is_ram) and (not is_flash) and (not is_pcb) and (
            cat_stripped == "cpu"
            or (not _has_cat and any(
                x in cat_desc
                for x in ("processor", "soc", "cpu", "microprocessor", "application processor")
            ))
        )
        is_ic = (not is_ram) and (not is_flash) and (not is_cpu) and (not is_pcb) and (
            _act_ic or (not _has_cat and any(
                x in cat_desc
                for x in (
                    "integrated circuit",
                    "controller",
                    "transceiver",
                    "semiconductor",
                    "amplifier",
                    "ic ",
                    " ic",
                )
            ))
        )
        # Tag-based IC fallback: U* tags without a category are treated as ICs
        if not any([is_ram, is_flash, is_cpu, is_ic, is_pcb]) and tags and tags[0].startswith("U"):
            is_ic = True

        is_connector = cat_stripped == "connector" or (not _has_cat and any(
            x in cat for x in ("connector", "header", "port", "socket")
        ))
        is_peripheral = cat_stripped == "peripheral"
        is_capacitor = cat_stripped == "capacitor" or (not _has_cat and "capacitor" in cat_desc)
        is_resistor = cat_stripped == "resistor" or (not _has_cat and "resistor" in cat_desc)
        is_inductor = cat_stripped == "inductor" or (not _has_cat and "inductor" in cat_desc)
        is_diode = any(x in cat for x in ("glass_smd", "diode", "led", "tvs")) or (
            not _has_cat and any(x in cat_desc for x in ("diode", "led", "tvs"))
        )

        # ── Silicon ──────────────────────────────────────────────────────────────
        if is_ram:
            key = _ic_key("memory" if tags else "dram", tags, value, part_slug)
            spec: dict[str, Any] = {"model": "dram", "n_ics": qty, "fab_yield": 0.875}
            if capacity_gb is not None:
                spec["capacity"] = f"{int(capacity_gb)} GB"
            spec["process"] = (
                _convert_process_node(mem_process) if mem_process else "ddr3_50nm"
            )
            bom["silicon"][key] = spec
            continue

        if is_flash:
            key = _ic_key("ssd", tags, value, part_slug)
            spec = {"model": "flash", "n_ics": qty, "fab_yield": 0.875}
            if capacity_gb is not None:
                spec["capacity"] = f"{int(capacity_gb)} GB"
            spec["process"] = (
                _convert_process_node(mem_process) if mem_process else "nand_30nm"
            )
            bom["silicon"][key] = spec
            continue

        if is_cpu:
            key = _ic_key("cpu", tags, value, part_slug)
            spec = {"n_ics": qty, **ic_defaults}
            if ic_die_mm2 is not None:
                area = _area_str(ic_die_mm2)
                if area:
                    spec["area"] = area
            spec["process"] = (
                _convert_process_node(soc_process_raw)
                if not _is_missing(soc_process_raw)
                else "28nm"
            )
            bom["silicon"][key] = spec
            continue

        if is_ic:
            key = _ic_key("ics", tags, value, part_slug)
            spec = {"n_ics": qty, **ic_defaults}
            if ic_die_mm2 is not None:
                area = _area_str(ic_die_mm2)
                if area:
                    spec["area"] = area
            spec["process"] = (
                _convert_process_node(soc_process_raw)
                if not _is_missing(soc_process_raw)
                else "45nm"
            )
            bom["silicon"][key] = spec
            continue

        # ── Materials ────────────────────────────────────────────────────────────
        if is_pcb:
            key = "pcb" if "pcb" not in bom["materials"] else f"pcb_{next_id('pcb')}"
            spec = {"category": "pcb"}
            if area_sq_mm is not None and area_sq_mm > 0:
                spec["area"] = f"{area_sq_mm / 100:.4f} cm2"
            if layers is not None:
                spec["layers"] = int(layers)
            if thickness_mm is not None:
                spec["thickness"] = f"{thickness_mm:.6g} mm"
            if not any(k for k in spec if k != "category"):
                spec["quantity"] = qty
            bom["materials"][key] = spec
            continue

        if is_solder:
            is_pb_free = any(
                x in cat_desc for x in ("pb-free", "pb free", "lead-free", "lead free", "pbfree", "rohs")
            )
            if is_pb_free:
                key = "solder.pbfree" if "solder.pbfree" not in bom["materials"] else f"solder.pbfree{next_id('solder')}"
                spec = {"category": "pb_free_solder", "type": "pb_free_solder"}
            else:
                key = "solder" if "solder" not in bom["materials"] else f"solder.tin{next_id('solder')}"
                spec = {"category": "tin", "type": "tin"}
            if weight_kg is not None and weight_kg > 0:
                spec["weight"] = _fmt_weight(weight_kg)
            else:
                spec["quantity"] = qty
            bom["materials"][key] = spec
            continue

        if is_shield:
            key = f"shield.metal{next_id('shield')}"
            spec = {"category": "tin", "type": "tin"}
            if weight_kg is not None and weight_kg > 0:
                spec["weight"] = _fmt_weight(weight_kg)
            else:
                spec["quantity"] = qty
            bom["materials"][key] = spec
            continue

        if is_aluminum:
            key = f"aluminum.alu{next_id('alu')}" if "aluminum" in bom["materials"] else "aluminum"
            spec = {"category": "aluminum", "type": "aluminum"}
            if weight_kg is not None and weight_kg > 0:
                spec["weight"] = _fmt_weight(weight_kg)
            else:
                spec["quantity"] = qty
            bom["materials"][key] = spec
            continue

        # ── Passives ─────────────────────────────────────────────────────────────
        if is_connector:
            key = _passive_key("connector", "con", tags)
            _peripheral_kw = (
                "usb", "hdmi", "ethernet", "magjack", "rj45",
                "headphone", "audio", "trs", "sd card", "microsd", "micro sd", "usd",
                "displayport", "sata", "thunderbolt",
            )
            _mfr_cat_lower = (e.get("category_mfr") or "").lower()
            _is_peripheral_conn = any(
                kw in desc_lower or kw in _mfr_cat_lower
                for kw in _peripheral_kw
            )
            conn_type = "peripheral" if _is_peripheral_conn else "generic"
            spec = {"category": "connector", "type": conn_type, "quantity": qty}
            if weight_kg is not None and weight_kg > 0:
                spec["weight"] = _fmt_weight(weight_kg)
            bom["passives"][key] = spec
            continue

        if is_peripheral:
            key = _passive_key("connector", "con", tags)
            spec = {"category": "connector", "type": "peripheral", "quantity": qty}
            if weight_kg is not None and weight_kg > 0:
                spec["weight"] = _fmt_weight(weight_kg)
            bom["passives"][key] = spec
            continue

        if is_capacitor:
            key = _passive_key("capacitor", "cap", tags)
            pkg = _extract_package_size(e.get("packaging") or e.get("type") or package_size_mm, desc)
            bom["passives"][key] = {"category": "capacitor", "type": pkg, "quantity": qty}
            continue

        if is_resistor:
            key = _passive_key("resistor", "res", tags)
            pkg = _extract_package_size(e.get("packaging") or e.get("type") or package_size_mm, desc)
            bom["passives"][key] = {"category": "resistor", "type": pkg, "quantity": qty}
            continue

        if is_inductor:
            key = _passive_key("inductor", "ind", tags)
            spec = {"category": "inductor", "quantity": qty}
            if weight_kg is not None and weight_kg > 0:
                spec["type"] = "weight_based"
                spec["weight"] = _fmt_weight(weight_kg)
            else:
                spec["type"] = _extract_package_size(e.get("packaging") or package_size_mm, desc)
            bom["passives"][key] = spec
            continue

        if is_diode:
            key = _passive_key("diode", "dio", tags)
            dtype = "led" if "led" in cat else "glass_smd"
            spec = {"category": "diode", "type": dtype, "quantity": qty}
            if weight_kg is not None and weight_kg > 0:
                spec["weight"] = _fmt_weight(weight_kg)
            bom["passives"][key] = spec
            continue

        # ── Other (explicit ACT category "other") ────────────────────────────────
        if is_other:
            type_val = (e.get("type") or "").strip()
            part = part_slug or f"other_{next_id('other')}"
            key = f"other.{part}"
            if key in bom["passives"]:
                key = f"other.{part}_{next_id('other')}"
            spec = {"category": "other", "quantity": qty}
            if type_val:
                spec["type"] = type_val
            if weight_kg is not None and weight_kg > 0:
                spec["weight"] = _fmt_weight(weight_kg)
            bom["passives"][key] = spec
            continue

        # ── Fallback ─────────────────────────────────────────────────────────────
        if tags:
            base_tag = tags[0]
            key_sfx = _key_suffix(value)
            key = f"{base_tag}.{key_sfx}" if key_sfx else base_tag
            if base_tag.startswith("U"):
                spec = {
                    "process": (
                        _convert_process_node(soc_process_raw)
                        if not _is_missing(soc_process_raw)
                        else "45nm"
                    ),
                    "n_ics": qty,
                    **ic_defaults,
                }
                if ic_die_mm2 is not None:
                    area = _area_str(ic_die_mm2)
                    if area:
                        spec["area"] = area
                bom["silicon"][key] = spec
            else:
                bom["passives"][key] = {"category": "other", "quantity": qty}
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