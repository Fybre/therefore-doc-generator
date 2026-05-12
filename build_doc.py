#!/usr/bin/env python3
"""
build_doc.py  -  Generate a Therefore implementation Word document.

Reads a Therefore XML configuration export, renders category index form
images inline, and produces a single .docx covering categories, indexing
profiles, workflows, folders, queries, reports, stamps, and keyword
dictionaries.

Usage:
    python build_doc.py TheConfiguration.xml [options]

    -o / --output     Output path (default: Therefore_Documentation.docx)
    --render-dir      Cache rendered PNGs here instead of a temp dir
    --data-dir        REST-pulled JSON dir for additional enrichment
    --no-images       Skip image rendering
    --skip-eforms     Exclude EForm-type indexing profiles
"""

import argparse
import json
import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

from themes import DEFAULT_THEME, Theme, load_theme

# Module-level theme — set by generate() before building.  All helpers read
# from this variable so we don't have to change every function signature.
_theme: Theme = DEFAULT_THEME

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Correct TypeNo mapping sourced from Therefore DB Codes reference.
# TypeNo 1-13 are fixed control types; negative values reference keyword
# dictionaries (via SingleTypeNo / MulitTypeNo) or data types.
FIELD_TYPE_NAMES = {
    1:  "String",
    2:  "Integer",
    3:  "Date",
    4:  "Label",
    5:  "Decimal",
    6:  "Logical",
    7:  "Datetime",
    8:  "Counter (integer)",
    9:  "Counter (formatted)",
    10: "Table field",
    11: "Integer (64-bit)",
    12: "Image",
    13: "Tab",
}

# Skip these types from the field table — visual-only or container elements.
SKIP_TYPES = {4, 13}

PROFILE_TYPE_LABELS = {1: "Import / capture", 2: "Workflow / case", 14: "EForm"}
SCRIPT_LANG_LABELS  = {1: "VBScript", 2: "C#", 3: "JavaScript"}

INDEXING_ERROR_MODE = {1: "Abort", 2: "Ignore", 3: "Use fallback", 4: "Report error"}
INDEXING_APPEND_MODE = {
    0: "Category default", 1: "Insert", 2: "Append", 3: "Replace",
    4: "No check", 5: "Error", 6: "Skip document",
}


def show_id(n) -> str:
    """Display a Therefore ID as its absolute value.
    Negative IDs are design-time identifiers (not yet committed to the DB);
    they're always shown without the minus sign, matching the HTML generator."""
    try:
        return str(abs(int(n)))
    except (TypeError, ValueError):
        return str(n) if n is not None else ""

# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------
def get_text(elem, path):
    e = elem.find(path)
    return e.text if e is not None else None


def get_int(elem, path, default=0):
    e = elem.find(path)
    if e is None or not e.text:
        return default
    try:
        return int(e.text)
    except ValueError:
        return default


def get_localised(elem, tag):
    """Extract localised string from <Tag><TStr><T><S>text</S></T></TStr></Tag>."""
    el = elem.find(tag)
    if el is None:
        return None
    s = el.find(".//TStr/T/S")
    if s is not None:
        return s.text
    s = el.find(".//S")
    return s.text if s is not None else None


def get_name(elem):
    n = get_localised(elem, "Name")
    if n:
        return n
    return get_text(elem, "Name") or get_text(elem, "Title")


def get_caption(field):
    cap = get_localised(field, "Caption")
    return cap or get_text(field, "ColName") or ""


def get_mandatory(field):
    """Try IsMandatory (correct element name) then Mandatory as fallback."""
    for tag in ("IsMandatory", "Mandatory"):
        v = get_text(field, tag)
        if v is not None:
            return v in ("1", "true", "True", "yes")
    return False


def get_regex(field):
    """Try RegEx (correct element name) then RegularExpr as fallback."""
    for tag in ("RegEx", "RegularExpr"):
        v = get_text(field, tag)
        if v:
            return v
    return None


def safe_fn(name):
    return re.sub(r"[^a-zA-Z0-9_-]", "_", name)


# ---------------------------------------------------------------------------
# Lookup maps  (built once from the full XML, passed around as a dict)
# ---------------------------------------------------------------------------
def parse_lookup_maps(root):
    """
    Parse all cross-reference tables from the XML.

    Returns a dict with keys:
        folder_paths   : {folder_no: "Parent / Child"} full path strings
        type_to_dict   : {type_no: {"name", "dict_no", "is_multi", "keywords"}}
        type_to_datatype: {type_no: "DataTypeName"}
        counter_map    : {counter_no: "CounterName"}
        kw_dicts       : [{dict_no, name, single_type_no, multi_type_no, keywords}]
        data_types     : [{dt_no, name, columns: [{col_no, col_name, caption, type_no}]}]
    """
    maps = {}

    # ── Folder paths ──────────────────────────────────────────────────────
    folder_nodes = {}   # folder_no → (name, parent_no)
    _fldrs = root.find("Folders")
    for fld in (_fldrs if _fldrs is not None else []):
        fno = get_int(fld, "FolderNo")
        if not fno:
            continue
        name = (get_localised(fld, "Name") or get_text(fld, "Name")
                or f"Folder_{fno}")
        # The processor tries <Parent> first, then <ParentNo>
        parent_el = fld.find("Parent")
        if parent_el is not None:
            try:
                parent_no = int(parent_el.text or 0)
            except ValueError:
                parent_no = 0
        else:
            parent_no = get_int(fld, "ParentNo") or get_int(fld, "ParentFolderNo")
        folder_nodes[fno] = (name, parent_no if parent_no else None)

    def _folder_path(fno, seen=None):
        if seen is None:
            seen = set()
        if fno not in folder_nodes or fno in seen:
            return str(fno)
        seen.add(fno)
        name, parent = folder_nodes[fno]
        if parent:
            return _folder_path(parent, seen) + " / " + name
        return name

    maps["folder_paths"] = {fno: _folder_path(fno) for fno in folder_nodes}

    # folder_nodes exposed for tree rendering: {folder_no → (name, parent_no)}
    maps["folder_nodes"] = folder_nodes
    # folder_children: {parent_no_or_0 → [child_folder_nos]} for recursive rendering
    folder_children = {}
    for fno, (name, parent) in folder_nodes.items():
        key = parent if parent else 0
        folder_children.setdefault(key, []).append(fno)
    maps["folder_children"] = folder_children

    # ── Keyword dictionaries ──────────────────────────────────────────────
    type_to_dict = {}
    kw_dicts = []
    _kd_el = root.find("KeywordDictionaries")
    for d in (_kd_el if _kd_el is not None else []):
        dict_no    = get_int(d, "KeyDicNo") or get_int(d, "DictionaryNo")
        name       = (get_text(d, "KeyDicName")
                      or get_localised(d, "Name")
                      or f"Dictionary_{dict_no}")
        single_tno = get_int(d, "SingleTypeNo") or None
        multi_tno  = get_int(d, "MulitTypeNo")  or None   # Note: "Mulit" is Therefore's typo

        keywords = []
        _kws_el = d.find("Keywords")
        for kw in (_kws_el if _kws_el is not None else []):
            kw_no  = get_int(kw, "KeywordNo") or get_int(kw, "KWNo")
            kw_val = (get_localised(kw, "Keyword")
                      or get_text(kw, "KWValue")
                      or get_text(kw, "Value") or "")
            if kw_val:
                keywords.append((kw_no, kw_val))

        entry = {
            "dict_no":        dict_no,
            "name":           name,
            "folder_no":      get_int(d, "FolderNo") or None,
            "single_type_no": single_tno,
            "multi_type_no":  multi_tno,
            "keywords":       keywords,
        }
        kw_dicts.append(entry)
        if single_tno:
            type_to_dict[single_tno] = {**entry, "is_multi": False}
        if multi_tno:
            type_to_dict[multi_tno]  = {**entry, "is_multi": True}

    maps["type_to_dict"] = type_to_dict
    maps["kw_dicts"]     = kw_dicts

    # ── Data types (referenced tables) ───────────────────────────────────
    type_to_datatype = {}
    data_types = []
    _dt_el = root.find("Datatypes")
    if _dt_el is None:
        _dt_el = root.find("DataTypes")
    for dt in (_dt_el if _dt_el is not None else []):
        dt_no   = get_int(dt, "TypeNo") or get_int(dt, "DataTypeNo")
        dt_name = get_text(dt, "Name") or f"DataType_{dt_no}"
        if dt_no:
            type_to_datatype[dt_no] = dt_name
        columns = []
        _cols_el = dt.find("Columns")
        for col in (_cols_el if _cols_el is not None else []):
            columns.append({
                "col_no":   get_int(col, "ColNo"),
                "col_name": get_text(col, "ColName") or "",
                "caption":  get_text(col, "Caption") or "",
                "type_no":  get_int(col, "TypeNo"),
                "length":   get_int(col, "Length"),
                "is_primary": get_text(col, "IsPrimary") in ("1", "true"),
            })
        data_types.append({
            "dt_no":     dt_no,
            "name":      dt_name,
            "folder_no": get_int(dt, "FolderNo") or None,
            "columns":   columns,
        })

    maps["type_to_datatype"] = type_to_datatype
    maps["data_types"]       = data_types

    # ── Counters ─────────────────────────────────────────────────────────
    counter_map = {}
    counters    = []
    _ctrs_el = root.find("Counters")
    for ctr in (_ctrs_el if _ctrs_el is not None else []):
        cno   = get_int(ctr, "CNo") or get_int(ctr, "CounterNo")
        cname = get_text(ctr, "Name") or f"Counter_{cno}"
        if cno:
            counter_map[cno] = cname
            counters.append({
                "counter_no": cno,
                "name":       cname,
                "folder_no":  get_int(ctr, "FolderNo") or None,
            })

    maps["counter_map"] = counter_map
    maps["counters"]    = counters

    # ── Field number → name map (for resolving [-370] in conditions) ─────
    field_no_map = {}
    _cats_el = root.find("Categories")
    for cat in (_cats_el if _cats_el is not None else []):
        _flds_el = cat.find("Fields")
        for fld in (_flds_el if _flds_el is not None else []):
            fno = get_text(fld, "FieldNo")
            if not fno:
                continue
            cap_el = fld.find("Caption")
            cap = ""
            if cap_el is not None:
                s = cap_el.find(".//S")
                cap = (s.text if s is not None else cap_el.text) or ""
            name = cap.strip() or get_text(fld, "ColName") or ""
            if name:
                field_no_map[fno] = name

    maps["field_no_map"] = field_no_map

    return maps


# ---------------------------------------------------------------------------
# Field reference resolution  [-370] → [FieldName]
# ---------------------------------------------------------------------------
_FIELD_REF_RE = re.compile(r'\[(-?\d+)\]')

def resolve_field_refs(text, field_no_map):
    """Replace [-370] style field number references with [FieldName]."""
    if not text or not field_no_map:
        return text
    return _FIELD_REF_RE.sub(
        lambda m: f"[{field_no_map.get(m.group(1), m.group(1))}]",
        text
    )


# ---------------------------------------------------------------------------
# Field type / detail resolution
# ---------------------------------------------------------------------------
def resolve_type(type_no, maps):
    """Return human-readable type name, resolving keyword/datatype refs."""
    td  = maps["type_to_dict"]
    tdt = maps["type_to_datatype"]
    if type_no is None:
        return "String"
    try:
        n = int(type_no)
    except (ValueError, TypeError):
        return str(type_no)

    if 1 <= n <= 13:
        return FIELD_TYPE_NAMES.get(n, f"Type {n}")
    if n < 0:
        if n in td:
            return "Keyword (multi)" if td[n]["is_multi"] else "Keyword (single)"
        if n in tdt:
            return "Referenced Table"
        return f"Type {n}"
    # n > 13 — might be a large datatype reference
    if n in tdt:
        return "Referenced Table"
    return f"Type {n}"


def resolve_type_detail(type_no, counter_no, maps):
    """Return type-specific detail strings (keyword list name, counter name, etc.)."""
    td      = maps["type_to_dict"]
    tdt     = maps["type_to_datatype"]
    ctr_map = maps["counter_map"]
    parts   = []
    try:
        n = int(type_no) if type_no is not None else 1
    except (ValueError, TypeError):
        return parts

    if n < 0:
        if n in td:
            d = td[n]
            parts.append(f"List: {d['name']}")
        elif n in tdt:
            parts.append(f"References: {tdt[n]}")
    elif n in (8, 9):
        try:
            cno = int(counter_no or 0)
        except ValueError:
            cno = 0
        if cno:
            parts.append(f"Counter: {ctr_map.get(cno, f'No {cno}')}")
    elif n > 13:
        if n in tdt:
            parts.append(f"References: {tdt[n]}")
    return parts


# ---------------------------------------------------------------------------
# Category parsing
# ---------------------------------------------------------------------------
def parse_categories(root):
    cats = []
    for cat in root.findall(".//Category"):
        n = get_text(cat, "CtgryNo")
        if n is None:
            continue
        name     = get_name(cat) or f"Category_{n}"
        folder_no = get_int(cat, "FolderNo") or None
        cats.append((int(n), name, folder_no, cat))
    cats.sort(key=lambda x: x[0])
    return cats


def cat_has_tabs(cat_elem):
    fields = cat_elem.find("Fields")
    if fields is None:
        return False
    return any(int(get_text(f, "TypeNo") or 1) == 13 for f in fields.findall("Field"))


def tab_names_for_cat(fields_elem):
    """Visible tabs only (used for field table's Tab: detail column)."""
    names = []
    for field in fields_elem.findall("Field"):
        if int(get_text(field, "TypeNo") or 1) != 13:
            continue
        tabs_el = field.find(".//Tabs")
        if tabs_el is None:
            continue
        for t in sorted(tabs_el.findall("T"), key=lambda x: int(get_text(x, "TabPos") or 999)):
            if get_text(t, "Visible") == "0":
                continue
            tc = t.find("TabCapt")
            s  = tc.find(".//S") if tc is not None else None
            if s is not None and s.text:
                names.append(s.text)
    return names


def _all_tab_names(fields_elem):
    """All tabs including system (Visible=0) — matches renderer output."""
    names = []
    for field in fields_elem.findall("Field"):
        if int(get_text(field, "TypeNo") or 1) != 13:
            continue
        tabs_el = field.find(".//Tabs")
        if tabs_el is None:
            continue
        for t in sorted(tabs_el.findall("T"), key=lambda x: int(get_text(x, "TabPos") or 999)):
            tc = t.find("TabCapt")
            s  = tc.find(".//S") if tc is not None else None
            if s is not None and s.text:
                names.append(s.text)
    return names


def field_tab_name(field, fields_elem):
    tab_no = get_text(field, "ShowInTabNo")
    if tab_no is None:
        return None
    for f in fields_elem.findall("Field"):
        if int(get_text(f, "TypeNo") or 1) != 13:
            continue
        tabs_el = f.find(".//Tabs")
        if tabs_el is None:
            continue
        for t in tabs_el.findall("T"):
            if get_text(t, "TabNo") == tab_no:
                tc = t.find("TabCapt")
                s  = tc.find(".//S") if tc is not None else None
                return s.text if s is not None else None
    return None


def category_fields(cat_elem, maps):
    """Return list of field dicts for the fields table, sorted by display order."""
    fields_elem = cat_elem.find("Fields")
    if fields_elem is None:
        return []

    rows = []
    for field in fields_elem.findall("Field"):
        if get_text(field, "Visible") == "0":
            continue
        type_no_str = get_text(field, "TypeNo")
        try:
            tno = int(type_no_str or 1)
        except ValueError:
            tno = 1
        if tno in SKIP_TYPES:
            continue

        caption    = get_caption(field)
        col_name   = get_text(field, "ColName") or ""
        field_no   = get_text(field, "FieldNo") or ""
        counter_no = get_text(field, "CounterNo")
        length     = get_text(field, "Length") or ""
        mandatory  = get_mandatory(field)
        default_v  = get_text(field, "DefaultVal") or ""
        regex      = get_regex(field)
        regex_help = (get_text(field, "RegExHelp") or "").strip()
        formula    = (get_text(field, "Formula") or "").strip()
        tab_name   = field_tab_name(field, fields_elem)
        index_type = get_int(field, "IndexType")

        display_name = f"{caption} ({col_name})" if col_name and col_name != caption else caption

        # Build details list: type-specific info first, then field properties
        details = resolve_type_detail(type_no_str, counter_no, maps)
        if tab_name:
            details.append(f"Tab: {tab_name}")
        if mandatory:
            details.append("Mandatory")
        if index_type == 2:
            details.append("Unique index")
        elif index_type == 1:
            details.append("Indexed")
        if default_v:
            details.append(f"Default: {default_v}")
        if formula:
            details.append(f"Formula: {formula}")
        if regex:
            details.append(f"Regex: {regex}")
        if regex_help:
            details.append(f"Regex help: {regex_help}")

        rows.append({
            "field_no":   show_id(field_no),
            "name":       display_name,
            "type":       resolve_type(type_no_str, maps),
            "size":       length,
            "details":    "; ".join(details),
            "disp_order": get_int(field, "DispOrderPos"),
            "tab_order":  get_int(field, "TabOrderPos"),
        })

    rows.sort(key=lambda r: (r["disp_order"], r["tab_order"]))
    return rows


# ---------------------------------------------------------------------------
# Indexing-profile parsing
# ---------------------------------------------------------------------------
def parse_ix_profiles(root, maps, include_eforms=True):
    cat_names  = {int(n): (get_name(cat) or f"Cat_{n}")
                  for cat in root.findall(".//Category")
                  if (n := get_text(cat, "CtgryNo")) is not None}

    field_lookup = {}
    for cat in root.findall(".//Category"):
        fe = cat.find("Fields")
        if fe is None:
            continue
        for f in fe.findall("Field"):
            fn = get_text(f, "FieldNo")
            if fn:
                cap = get_localised(f, "Caption") or get_text(f, "ColName")
                field_lookup[int(fn)] = cap or f"Field {fn}"

    profiles = []
    _ix_el = root.find("IxProfiles")
    for p in (_ix_el if _ix_el is not None else []):
        ptype = int(get_text(p, "ProfileType") or 0)
        if not include_eforms and ptype == 14:
            continue

        folder_no  = int(get_text(p, "FolderNo") or 0)
        def_id     = p.find("DefId")
        src_id     = p.find("SourceDefId")
        tgt_no     = int(def_id.findtext("DefNo")) if def_id is not None else 0
        src_no     = int(src_id.findtext("DefNo")) if src_id is not None else 0
        folder_path = maps["folder_paths"].get(folder_no, "(root)") if folder_no else "(root)"

        assignments = []
        _asgn_el = p.find("Assignments")
        for a in (_asgn_el if _asgn_el is not None else []):
            fn  = int(get_text(a, "FieldNo") or 0)
            em  = int(get_text(a, "ErrorMode") or 0)
            assignments.append({
                "field":      field_lookup.get(fn, f"Field {fn}"),
                "expr":       (get_text(a, "Expression")     or "").strip(),
                "check":      (get_text(a, "CheckCondition") or "").strip(),
                "fallback":   (get_text(a, "ErrorFallback")  or "").strip(),
                "error_mode": INDEXING_ERROR_MODE.get(em, str(em)) if em else "",
            })

        aa = int(get_text(p, "AutoAppendMode") or 0)
        profiles.append({
            "name":         get_text(p, "Name") or "",
            "ptype":        ptype,
            "ptype_label":  PROFILE_TYPE_LABELS.get(ptype, f"Type {ptype}"),
            "lang":         int(get_text(p, "ScriptLang") or 0),
            "folder_name":  folder_path,
            "init_script":  (get_text(p, "InitScript") or "").strip(),
            "filter":       (get_text(p, "Filter") or "").strip(),
            "auto_append":  INDEXING_APPEND_MODE.get(aa, str(aa)) if aa else "",
            "target_cat":   cat_names.get(tgt_no, ""),
            "source_cat":   cat_names.get(src_no, "") if src_no else "",
            "assignments":  assignments,
        })

    return profiles


# ---------------------------------------------------------------------------
# Document formatting helpers
# ---------------------------------------------------------------------------
def _set_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    tcPr.append(parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{hex_color}"/>'
    ))


def _set_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcPr.append(parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top    w:val="single" w:sz="4" w:space="0" w:color="{_theme.hex("table_border")}"/>'
        f'<w:left   w:val="single" w:sz="4" w:space="0" w:color="{_theme.hex("table_border")}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{_theme.hex("table_border")}"/>'
        f'<w:right  w:val="single" w:sz="4" w:space="0" w:color="{_theme.hex("table_border")}"/>'
        '</w:tcBorders>'
    ))


def hdr_cell(cells, idx, text):
    c = cells[idx]
    c.text = ""
    _set_bg(c, _theme.hex("primary"))
    _set_borders(c)
    run = c.paragraphs[0].add_run(text)
    run.bold = True
    run.font.color.rgb = _theme.docx_rgb("white")
    run.font.size = Pt(9)


def body_cell(cells, idx, text, fill, bold=False, mono=False, pt=9):
    c = cells[idx]
    c.text = ""
    _set_bg(c, fill)
    _set_borders(c)
    run = c.paragraphs[0].add_run(str(text or ""))
    run.bold = bold
    run.font.size = Pt(pt)
    if mono:
        run.font.name = "Consolas"


def _set_para_shading(para, hex_fill):
    from docx.oxml import OxmlElement
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    pPr.append(shd)


def _set_para_border(para, hex_color=None, side="left", sz="12", space="72"):
    if hex_color is None:
        hex_color = _theme.hex("code_border")
    from docx.oxml import OxmlElement
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bdr = OxmlElement(f"w:{side}")
    bdr.set(qn("w:val"), "single")
    bdr.set(qn("w:sz"), sz)
    bdr.set(qn("w:space"), space)
    bdr.set(qn("w:color"), hex_color)
    pBdr.append(bdr)
    pPr.append(pBdr)


def add_code_block(doc, text, lang_label=None):
    """Render a multi-line code string as a shaded monospace block."""
    if lang_label:
        lp = doc.add_paragraph()
        lr = lp.add_run(lang_label)
        lr.bold = True
        lr.font.size = Pt(8)
        lr.font.color.rgb = _theme.docx_rgb("muted")
        lp.paragraph_format.space_after = Pt(0)

    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        p.paragraph_format.left_indent  = Cm(0.4)
        _set_para_shading(p, _theme.hex("code_bg"))
        _set_para_border(p, side="left", sz="16", space="72")
        run = p.add_run(line if line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(8)
    doc.add_paragraph().paragraph_format.space_before = Pt(4)


RENDER_NOTE = "Note: this render is indicative only and may not exactly match the appearance in Therefore."

def add_render_note(doc):
    p = doc.add_paragraph(RENDER_NOTE)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.runs[0]
    r.font.size = Pt(8)
    r.font.italic = True
    r.font.color.rgb = _theme.docx_rgb("light_muted")
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(6)


def set_table_style(table):
    """Set table to full page width, disable autofit, and apply compact cell margins."""
    from docx.oxml import OxmlElement
    tbl   = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    # Full width (A4 minus 1.8cm each side = ~17.4cm = 9866 twips)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    "9866")
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)

    # Disable autofit so our explicit column widths are respected
    tblLook = OxmlElement("w:tblLayout")
    tblLook.set(qn("w:type"), "fixed")
    tblPr.append(tblLook)

    # Compact cell margins
    mar = OxmlElement("w:tblCellMar")
    for side, val in [("top", 30), ("bottom", 30), ("left", 60), ("right", 60)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def set_col_widths(table, widths_cm):
    set_table_style(table)
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


def _manual_heading(doc, text, size_pt, space_before, space_after,
                    meta="", bottom_border=False):
    """
    Add a heading as a plain Normal paragraph with manual font/colour styling.
    Avoids Word's Heading styles so wrapper-document auto-numbering never fires.
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before  = Pt(space_before)
    p.paragraph_format.space_after   = Pt(space_after)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold           = True
    r.font.name      = _theme.font("body")
    r.font.size      = Pt(size_pt)
    r.font.color.rgb = _theme.docx_rgb("primary")
    if meta:
        m = p.add_run(f"   {meta}")
        m.font.name      = _theme.font("body")
        m.font.size      = Pt(size_pt - 2)
        m.font.color.rgb = _theme.docx_rgb("muted")
        m.font.bold      = False
    if bottom_border:
        _set_para_border(p, hex_color=_theme.hex("h1_border"), side="bottom", sz="6", space="1")
    return p


def blue_heading(doc, text, level, meta=""):
    """Wrapper kept for call-site compatibility."""
    if level == 1:
        return _manual_heading(doc, text, 15, 16, 5, meta=meta, bottom_border=True)
    elif level == 2:
        return _manual_heading(doc, text, 12, 12, 3, meta=meta)
    else:
        return _manual_heading(doc, text, 10,  8, 2, meta=meta)


def kv_table(doc, rows, lw=5.0, vw=11.5):
    """Two-column label/value metadata table; skips empty values."""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    for label, value in rows:
        if not value:
            continue
        row = t.add_row()
        body_cell(row.cells, 0, label, _theme.hex("grey"), bold=True)
        body_cell(row.cells, 1, value, _theme.hex("white"))
    set_col_widths(t, [lw, vw])
    return t


def sec_heading(section_no, title):
    """Format a section heading — omits the number when section_no is None (body-only mode)."""
    return f"{section_no}. {title}" if section_no else title


# ---------------------------------------------------------------------------
# Section: overview
# ---------------------------------------------------------------------------
def build_overview(doc, counts, section_no=1):
    _manual_heading(doc, sec_heading(section_no, "Implementation Overview"), 15, 16, 5, bottom_border=True)
    for label, count in counts:
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(str(count))


# ---------------------------------------------------------------------------
# Section: categories
# ---------------------------------------------------------------------------
def _embed_images(doc, name, cat_elem, render_dir, ext=".png"):
    ctgry_no = cat_elem.findtext("CtgryNo") or "x"
    if cat_has_tabs(cat_elem):
        fields_el = cat_elem.find("Fields")
        if fields_el is not None:
            # tab_names_for_cat only returns visible tabs — but render now includes
            # all tabs, so scan the render dir for any matching tab images
            for tab in _all_tab_names(fields_el):
                img = os.path.join(render_dir, f"{safe_fn(name)}_{ctgry_no}_tab_{safe_fn(tab)}{ext}")
                if os.path.exists(img):
                    p = doc.add_paragraph()
                    p.add_run(f"Tab: {tab}").bold = True
                    pp = doc.add_paragraph()
                    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    pp.add_run().add_picture(img, width=Cm(14.0))
    else:
        img = os.path.join(render_dir, f"{safe_fn(name)}_{ctgry_no}{ext}")
        if os.path.exists(img):
            pp = doc.add_paragraph()
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.add_run().add_picture(img, width=Cm(14.0))


def build_fields_table(doc, rows):
    if not rows:
        return
    doc.add_paragraph().add_run("Fields:").bold = True
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    for i, h in enumerate(["No", "Field Name", "Type", "Size", "Details"]):
        hdr_cell(hdr, i, h)
    for idx, fr in enumerate(rows):
        fill = _theme.hex("grey") if idx % 2 == 1 else _theme.hex("white")
        rc = t.add_row().cells
        body_cell(rc, 0, fr["field_no"], fill)
        body_cell(rc, 1, fr["name"],     fill, bold=True)
        body_cell(rc, 2, fr["type"],     fill)
        body_cell(rc, 3, fr["size"],     fill)
        body_cell(rc, 4, fr["details"],  fill)
    set_col_widths(t, [1.2, 5.8, 3.6, 1.3, 5.5])


def build_categories(doc, categories, render_dir, include_images, maps, section_no=2, img_format="png"):
    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Categories"), 15, 16, 5, bottom_border=True)
    doc.add_paragraph(
        "This section documents all category index forms, "
        "including visual renders and field definitions."
    )
    fp_map = maps["folder_paths"]
    # Word embedding always uses PNG (SVG fallbacks are generated when needed)
    for ctgry_no, name, folder_no, cat_elem in categories:
        meta = f"(Category No: {show_id(ctgry_no)})"
        blue_heading(doc, name, level=2, meta=meta)

        if folder_no and folder_no in fp_map:
            p = doc.add_paragraph()
            p.add_run("Folder: ").italic = True
            run = p.add_run(fp_map[folder_no])
            run.bold = True
            run.font.color.rgb = _theme.docx_rgb("primary")

        desc = (cat_elem.findtext("Description") or "").strip()
        if desc:
            dp = doc.add_paragraph(desc)
            dp.runs[0].font.color.rgb = _theme.docx_rgb("description")
            dp.runs[0].italic = True

        if include_images and render_dir:
            _embed_images(doc, name, cat_elem, render_dir)

        build_fields_table(doc, category_fields(cat_elem, maps))
        doc.add_paragraph()


# ---------------------------------------------------------------------------
# Section: indexing profiles
# ---------------------------------------------------------------------------
def build_ix_profiles(doc, profiles, section_no, maps=None):
    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Indexing Profiles"), 15, 16, 5, bottom_border=True)
    doc.add_paragraph(
        f"{len(profiles)} indexing profiles. Each defines how documents flow into a "
        "target category — filter, init script, and per-field extraction expressions."
    )

    by_folder = {}
    for p in profiles:
        by_folder.setdefault(p["folder_name"], []).append(p)

    FIRST = ["System Profiles", "MFiles Import Profiles", "(root)"]
    order = [k for k in FIRST if k in by_folder]
    order += [k for k in by_folder if k not in FIRST]

    for folder in order:
        fps = by_folder[folder]
        display = "Root (no folder)" if folder == "(root)" else folder
        blue_heading(doc, display, level=2,
                     meta=f"({len(fps)} profile{'s' if len(fps) != 1 else ''})")

        for p in sorted(fps, key=lambda x: x["name"]):
            blue_heading(doc, p["name"], level=3)

            fno_map = (maps or {}).get("field_no_map", {})
            lang_label = SCRIPT_LANG_LABELS.get(p["lang"], "") if p["lang"] else ""
            kv_table(doc, [
                ("Target Category", p["target_cat"]),
                ("Source Category", p["source_cat"]),
                ("Profile Type",    p["ptype_label"]),
                ("Script Language", lang_label),
                ("Auto-Append",     p["auto_append"]),
                ("Filter",          resolve_field_refs(p["filter"], fno_map)),
            ])
            doc.add_paragraph()

            if p["init_script"]:
                hp = doc.add_paragraph()
                r = hp.add_run("Init Script:")
                r.bold = True
                r.font.color.rgb = _theme.docx_rgb("primary")
                hp.paragraph_format.space_after = Pt(2)
                add_code_block(doc, p["init_script"])

            if p["assignments"]:
                ap = doc.add_paragraph()
                r = ap.add_run(f"Field assignments ({len(p['assignments'])}):")
                r.bold = True
                r.font.color.rgb = _theme.docx_rgb("primary")

                any_check    = any(a["check"]    for a in p["assignments"])
                any_fallback = any(a["fallback"] for a in p["assignments"])

                if any_check or any_fallback:
                    cols   = ["Field", "Expression", "Check Condition", "Error Mode"]
                    widths = [4.5, 7.5, 4.5, 1.5]
                else:
                    cols   = ["Field", "Expression", "Error Mode"]
                    widths = [5.0, 11.5, 1.8]

                t = doc.add_table(rows=1, cols=len(cols))
                t.style = "Table Grid"
                for i, h in enumerate(cols):
                    hdr_cell(t.rows[0].cells, i, h)
                for idx, a in enumerate(p["assignments"]):
                    fill = _theme.hex("grey") if idx % 2 == 1 else _theme.hex("white")
                    rc = t.add_row().cells
                    body_cell(rc, 0, a["field"],  fill, bold=True)
                    body_cell(rc, 1, a["expr"],   fill, mono=True, pt=8)
                    if any_check or any_fallback:
                        body_cell(rc, 2, a["check"],      fill, mono=True, pt=8)
                        body_cell(rc, 3, a["error_mode"], fill, pt=8)
                    else:
                        body_cell(rc, 2, a["error_mode"], fill, pt=8)
                set_col_widths(t, widths)
                doc.add_paragraph()


# ---------------------------------------------------------------------------
# Section: workflows
# ---------------------------------------------------------------------------
TASK_TYPE_LABELS = {"1": "Start", "2": "End", "3": "Manual", "4": "Automatic"}

try:
    from render_workflow import render_workflow as _render_workflow
    _WORKFLOW_RENDER_AVAILABLE = True
except ImportError:
    _WORKFLOW_RENDER_AVAILABLE = False


def build_workflows(doc, root, section_no, maps, render_dir=None, font=None):
    wf_parent = root.find("WFProcesses")
    if wf_parent is None:
        wf_parent = root.find("Workflows")
    if wf_parent is None:
        return
    workflows = wf_parent.findall("WFProcess") or wf_parent.findall("Workflow")
    if not workflows:
        return

    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Workflow Processes"), 15, 16, 5, bottom_border=True)
    doc.add_paragraph(f"Total workflow processes: {len(workflows)}")

    for wf in workflows:
        wf_no = get_text(wf, "ProcessNo") or get_text(wf, "WFNo")
        name  = get_name(wf) or f"Workflow_{wf_no}"
        blue_heading(doc, name, level=2, meta=f"(Process No: {show_id(wf_no)})")

        # Embed workflow diagram if renderer is available
        if _WORKFLOW_RENDER_AVAILABLE and render_dir and font:
            img_path = os.path.join(render_dir, f"{safe_fn(name)}_{wf_no}.png")
            result = _render_workflow(wf, img_path, font, field_no_map=maps.get("field_no_map", {}))
            if result and os.path.exists(result):
                from PIL import Image as _PilImg
                _iw, _ih = _PilImg.open(result).size
                max_w = Cm(16.5)
                max_h = Cm(12.7)
                if _iw > 0 and (_ih / _iw) * max_w > max_h:
                    pic_kwargs = {"height": max_h}
                else:
                    pic_kwargs = {"width": max_w}
                pp = doc.add_paragraph()
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pp.add_run().add_picture(result, **pic_kwargs)

        tasks_el = wf.find("Tasks")
        if tasks_el is None:
            continue
        task_list = tasks_el.findall("T") or tasks_el.findall("Task")
        if not task_list:
            continue
        task_list = sorted(task_list, key=lambda t: int(get_text(t, "SeqPos") or 999))

        # Build task_no → name map for transition target lookup
        task_name_map = {}
        for t in task_list:
            tno  = get_text(t, "TaskNo") or ""
            tnam = get_name(t) or f"Task {tno}"
            task_name_map[tno] = tnam

        # Tasks table: # | Name | Type | Assigned To | Duration | Transitions
        t = doc.add_table(rows=1, cols=6)
        t.style = "Table Grid"
        for i, h in enumerate(["#", "Task Name", "Type", "Assigned To", "Duration", "Transitions"]):
            hdr_cell(t.rows[0].cells, i, h)

        for idx, task in enumerate(task_list):
            fill       = _theme.hex("grey") if idx % 2 == 1 else _theme.hex("white")
            task_no    = get_text(task, "TaskNo") or ""
            task_name  = get_name(task) or f"Task_{task_no}"
            task_type  = get_text(task, "Type") or ""
            type_label = TASK_TYPE_LABELS.get(task_type, task_type or "—")
            seq_pos    = get_text(task, "SeqPos") or str(idx + 1)
            duration   = get_int(task, "Duration")
            dur_str    = f"{duration}h" if duration else "—"

            # Assigned users from Choices/CH
            assigned = []
            choices_el = task.find("Choices")
            if choices_el is not None:
                for ch in choices_el.findall("CH"):
                    display = (get_text(ch, "DisplayName") or get_text(ch, "UserName") or "")
                    if display:
                        assigned.append(display)
            assigned_str = ", ".join(assigned) if assigned else "—"

            # Transitions — collect structured list
            transitions = []
            trans_el = task.find("Transitions")
            if trans_el is not None:
                for tr in trans_el.findall("TR"):
                    to_no     = get_text(tr, "TaskToNo") or ""
                    to_name   = task_name_map.get(to_no, f"Task {to_no}")
                    action    = (get_localised(tr, "ActionText") or get_name(tr) or "").strip()
                    condition = resolve_field_refs(
                        (get_text(tr, "Condition") or "").strip(),
                        maps.get("field_no_map", {})
                    )
                    # Suppress action when it's the same as the target name
                    if action.lower() == to_name.lower():
                        action = ""
                    transitions.append((to_name, action, condition))

            rc = t.add_row().cells
            body_cell(rc, 0, show_id(seq_pos), fill)
            body_cell(rc, 1, task_name, fill, bold=True)
            body_cell(rc, 2, type_label, fill)
            body_cell(rc, 3, assigned_str, fill)
            body_cell(rc, 4, dur_str, fill)

            # Build transitions cell with formatted runs
            tc = rc[5]
            tc.text = ""
            _set_bg(tc, fill)
            _set_borders(tc)
            if not transitions:
                tc.paragraphs[0].add_run("—").font.size = Pt(8)
            else:
                for i, (to_name, action, condition) in enumerate(transitions):
                    p = tc.paragraphs[0] if i == 0 else tc.add_paragraph()
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after  = Pt(0)
                    arrow = p.add_run(f"→ {to_name}")
                    arrow.font.size = Pt(8)
                    if action:
                        act_r = p.add_run(f"  ({action})")
                        act_r.font.size = Pt(7)
                        act_r.font.color.rgb = _theme.docx_rgb("muted")
                    if condition:
                        cp = tc.add_paragraph()
                        cp.paragraph_format.space_before = Pt(0)
                        cp.paragraph_format.space_after  = Pt(2)
                        cp.paragraph_format.left_indent  = Cm(0.4)
                        cr = cp.add_run(f"IF: {condition}")
                        cr.font.size     = Pt(7)
                        cr.font.name     = "Consolas"
                        cr.font.color.rgb = _theme.docx_rgb("light_muted")
                        cr.font.italic   = True

        set_col_widths(t, [1.0, 4.5, 2.3, 3.8, 1.8, 4.0])
        doc.add_paragraph()


# ---------------------------------------------------------------------------
# Section: eForms
# ---------------------------------------------------------------------------
_EFORM_RENDER_AVAILABLE = False
_render_eform = None
try:
    from render_eform import render_eform as _render_eform
    _EFORM_RENDER_AVAILABLE = True
except ImportError:
    pass

EFORM_TYPE_LABELS = {
    "textfield": "Text", "textarea": "Textarea", "number": "Number",
    "select": "Select", "checkbox": "Checkbox", "datetime": "Date/Time",
    "email": "Email", "phoneNumber": "Phone", "lookup": "Lookup",
    "datagrid": "Data Grid", "signature": "Signature", "hidden": "Hidden",
    "radio": "Radio", "selectboxes": "Checkboxes",
    "htmlelement": "HTML", "content": "HTML",
}


def _collect_eform_fields(components, _page="", _datagrid=""):
    """Recursively flatten all user-facing fields from a Formio component tree."""
    rows = []
    for c in components:
        t = c.get("type", "")
        label = c.get("label", "") or c.get("title", "")

        def _has_script(comp):
            return any(comp.get(k) for k in
                       ("logic", "customConditional", "calculateValue",
                        "validate"))

        if t == "panel":
            page = label or _page
            if _has_script(c):
                pass  # fall through to add this panel as a row too
            else:
                rows += _collect_eform_fields(c.get("components") or [], _page=page,
                                              _datagrid=_datagrid)
                continue
        elif t == "columns":
            for col in (c.get("columns") or []):
                rows += _collect_eform_fields(col.get("components") or [],
                                              _page=_page, _datagrid=_datagrid)
            if not _has_script(c):
                continue
        elif t == "datagrid":
            rows += _collect_eform_fields(c.get("components") or [],
                                          _page=_page, _datagrid=label)
            if not _has_script(c):
                continue
        elif t == "button":
            if not _has_script(c):
                continue

        # For HTML elements, extract readable text from content as label
        if t in ("htmlelement", "content"):
            content = c.get("content", "")
            import html as _html, re as _re
            text = _re.sub(r"<style[^>]*>.*?</style>", " ", content, flags=_re.S)
            text = _re.sub(r"<script[^>]*>.*?</script>", " ", text, flags=_re.S)
            text = _re.sub(r"<[^>]+>", " ", text)
            text = _html.unescape(text)
            text = _re.sub(r"\s+", " ", text).strip()
            if text and not _re.match(r"^[\{\}\(\);:#\.\s]+$", text):
                label = text[:80]

        # Build validation string
        v       = c.get("validate") or {}
        req     = bool(v.get("required"))
        rules   = []
        if v.get("minLength"):  rules.append(f"Min length: {v['minLength']}")
        if v.get("maxLength"):  rules.append(f"Max length: {v['maxLength']}")
        if v.get("min") not in (None, ""):  rules.append(f"Min: {v['min']}")
        if v.get("max") not in (None, ""):  rules.append(f"Max: {v['max']}")
        if v.get("pattern"):    rules.append(f"Pattern: {v['pattern']}")
        if v.get("custom"):     rules.append(f"Custom: {v['custom'].strip()}")

        # Conditional visibility
        cond_str = ""
        cond = c.get("conditional") or {}
        cc   = (c.get("customConditional") or "").strip()
        if cc:
            cond_str = cc
        elif cond.get("when") and str(cond.get("show", "")).strip() not in ("", "null"):
            show = "show" if str(cond["show"]).lower() in ("true", "1") else "hide"
            cond_str = f"{show} when {cond['when']} = {cond.get('eq', '')}"
        elif cond.get("json") and cond["json"] not in ("", None, {}, []):
            try:
                cond_str = f"JSON: {json.dumps(cond['json'])[:80]}"
            except Exception:
                cond_str = str(cond["json"])[:80]

        # Calculated value
        calc = (c.get("calculateValue") or "").strip()

        # Default value
        dv = c.get("defaultValue")
        dv_str = ""
        if dv not in (None, "", False, 0, []):
            dv_str = str(dv)[:60]

        # Formio advanced logic
        logic_parts = []
        for l in (c.get("logic") or []):
            tr = l.get("trigger") or {}
            tr_type = tr.get("type", "")
            # Event triggers show the event name; JS/simple triggers show code
            if tr_type == "event":
                tr_desc = f"on event: {tr.get('event', '')}"
            else:
                tr_val = (tr.get("javascript") or tr.get("json") or tr.get("simple") or "")
                if isinstance(tr_val, dict):
                    tr_val = json.dumps(tr_val)
                tr_desc = f"{tr_type}: {str(tr_val).strip()}"
            acts = []
            for a in (l.get("actions") or []):
                atype = a.get("type", "")
                prop  = (a.get("property") or {}).get("value", "")
                val   = str(a.get("value", "") or a.get("customAction", "")).strip()
                if prop:
                    acts.append(f"set {prop} = {val}")
                else:
                    acts.append(f"{atype}: {val}")
            logic_parts.append(
                f"Trigger ({tr_desc})\n  → " + "\n  → ".join(acts)
                if acts else f"Trigger ({tr_desc})"
            )

        rows.append({
            "page":     _page,
            "datagrid": _datagrid,
            "label":    label,
            "type":     t,
            "key":      c.get("key", ""),
            "required": req,
            "disabled": bool(c.get("disabled")),
            "hidden":   bool(c.get("hidden")),
            "rules":    "\n".join(rules),
            "calc":     calc,
            "default":  dv_str,
            "cond":     cond_str,
            "logic":    "\n".join(logic_parts),
        })
    return rows


def _add_eform_fields_table(doc, ef_elem, maps=None):
    import json as _json
    fdef_text = ef_elem.findtext("FDef")
    if not fdef_text:
        return
    try:
        fdef = _json.loads(fdef_text)
    except Exception:
        return

    fno_map = (maps or {}).get("field_no_map", {})
    fields = _collect_eform_fields(fdef.get("components") or [])
    for f in fields:
        f["calc"]  = resolve_field_refs(f["calc"],  fno_map)
        f["cond"]  = resolve_field_refs(f["cond"],  fno_map)
        f["logic"] = resolve_field_refs(f["logic"], fno_map)
        f["rules"] = resolve_field_refs(f["rules"], fno_map)
    if not fields:
        return

    # Table columns: Page | Field | Type | Req | Validation / Script / Logic
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    for i, h in enumerate(["Page / Grid", "Field", "Type", "Req", "Validation / Script / Logic"]):
        hdr_cell(t.rows[0].cells, i, h)

    for idx, f in enumerate(fields):
        fill = _theme.hex("grey") if idx % 2 == 1 else _theme.hex("white")
        rc   = t.add_row().cells

        # Page / datagrid context
        ctx = f["page"]
        if f["datagrid"]:
            ctx = f"{ctx} › {f['datagrid']}" if ctx else f["datagrid"]
        body_cell(rc, 0, ctx, fill, pt=8)

        # Field label + key
        c0 = rc[1]
        c0.text = ""
        _set_bg(c0, fill)
        _set_borders(c0)
        r = c0.paragraphs[0].add_run(f["label"] or f["key"])
        r.bold = True
        r.font.size = Pt(9)
        if f["hidden"]:
            r.font.color.rgb = _theme.docx_rgb("footer_text")
        kr = c0.paragraphs[0].add_run(f"\n{f['key']}")
        kr.font.size = Pt(7)
        kr.font.color.rgb = _theme.docx_rgb("light_muted")
        kr.font.name = "Consolas"

        # Type + flags
        type_lbl = EFORM_TYPE_LABELS.get(f["type"], f["type"])
        flags = []
        if f["disabled"]: flags.append("read-only")
        if f["hidden"]:   flags.append("hidden")
        type_str = type_lbl + (f" ({', '.join(flags)})" if flags else "")
        body_cell(rc, 2, type_str, fill, pt=8)

        # Required
        body_cell(rc, 3, "✓" if f["required"] else "", fill, pt=9)

        # Validation / Script / Logic cell — build multi-run paragraph
        c4 = rc[4]
        c4.text = ""
        _set_bg(c4, fill)
        _set_borders(c4)
        para = c4.paragraphs[0]
        para.paragraph_format.space_after = Pt(0)

        def _script_run(p, label_text, code, mono=True):
            lr = p.add_run(label_text + ": ")
            lr.bold = True
            lr.font.size = Pt(8)
            lr.font.color.rgb = _theme.docx_rgb("primary")
            cr = p.add_run(code)
            cr.font.size = Pt(8)
            if mono:
                cr.font.name = "Consolas"
            cr.font.color.rgb = _theme.docx_rgb("text_dark")

        added = False
        if f["rules"]:
            _script_run(para, "Validation", f["rules"], mono=False)
            added = True
        if f["default"]:
            if added: para.add_run("\n")
            _script_run(para, "Default", f["default"], mono=True)
            added = True
        if f["cond"]:
            if added: para.add_run("\n")
            _script_run(para, "Condition", f["cond"], mono=True)
            added = True
        if f["calc"]:
            if added: para.add_run("\n")
            _script_run(para, "Calculated", f["calc"], mono=True)
            added = True
        if f["logic"]:
            if added: para.add_run("\n")
            _script_run(para, "Logic", f["logic"], mono=True)

    set_col_widths(t, [2.5, 3.5, 2.5, 1.0, 7.9])


def build_eforms(doc, root, section_no, maps=None, render_dir=None, font=None):
    ef_parent = root.find("EForms")
    if ef_parent is None:
        return
    eforms = list(ef_parent)
    if not eforms:
        return

    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "eForms"), 15, 16, 5, bottom_border=True)
    doc.add_paragraph(f"Total eForms: {len(eforms)}")

    for ef in eforms:
        name  = ef.findtext("FName") or "EForm"
        fno   = ef.findtext("FNo") or "x"
        blue_heading(doc, name, level=2, meta=f"(Form No: {show_id(fno)})")

        if _EFORM_RENDER_AVAILABLE and render_dir:
            img_path = os.path.join(render_dir, f"eform_{safe_fn(name)}_{fno}.png")
            result = _render_eform(ef, img_path, font)
            if result and os.path.exists(result):
                from PIL import Image as _PilImg
                _iw, _ih = _PilImg.open(result).size
                max_w = Cm(16.5)
                max_h = Cm(17.8)
                if _iw > 0 and (_ih / _iw) * max_w > max_h:
                    pic_kwargs = {"height": max_h}
                else:
                    pic_kwargs = {"width": max_w}
                pp = doc.add_paragraph()
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pp.add_run().add_picture(result, **pic_kwargs)

        _add_eform_fields_table(doc, ef, maps=maps)
        doc.add_paragraph()


# ---------------------------------------------------------------------------
# Section: folder / object hierarchy
# ---------------------------------------------------------------------------
def build_folder_hierarchy(doc, root, maps, categories, section_no):
    """Render a recursive folder tree with all object types nested inside."""
    fn_nodes  = maps["folder_nodes"]    # {folder_no → (name, parent_no)}
    fn_childs = maps["folder_children"] # {parent_no_or_0 → [folder_nos]}
    if not fn_nodes:
        return

    # Collect all objects by folder_no ─────────────────────────────────────
    # type_order controls display sort within each folder
    TYPE_ORDER = ["Category", "Workflow", "Query", "EForm", "Dictionary",
                  "Counter", "Stamp", "Data Type"]

    folder_items = {}  # {folder_no → [(type_label, name, obj_no)]}

    def add_item(folder_no, type_label, name, obj_no):
        key = folder_no if folder_no else 0
        folder_items.setdefault(key, []).append((type_label, name, obj_no))

    for ctgry_no, name, folder_no, _ in categories:
        add_item(folder_no, "Category", name, ctgry_no)

    wfp = root.find("WFProcesses")
    if wfp is None:
        wfp = root.find("Workflows")
    for wf in (wfp if wfp is not None else []):
        fn   = get_int(wf, "FolderNo") or None
        wno  = get_text(wf, "ProcessNo") or get_text(wf, "WFNo")
        wnam = get_name(wf) or f"Workflow_{wno}"
        add_item(fn, "Workflow", wnam, wno)

    qp = root.find("QueryTemplates")
    qtag = "QueryTemplate"
    qno_tag = "QueryTemplateNo"
    if qp is None:
        qp = root.find("Queries")
        qtag = "Query"
        qno_tag = "QueryNo"
    for q in (qp.findall(qtag) if qp is not None else []):
        fn  = get_int(q, "FolderNo") or None
        qno = get_text(q, qno_tag) or get_text(q, "QueryNo")
        qnam = get_name(q) or f"Query_{qno}"
        add_item(fn, "Query", qnam, qno)

    ef_parent = root.find("EForms")
    for ef in (ef_parent if ef_parent is not None else []):
        fn   = get_int(ef, "FFold") or None
        eno  = ef.findtext("FNo")
        enam = ef.findtext("FName") or f"EForm_{eno}"
        add_item(fn, "EForm", enam, eno)

    for d in maps["kw_dicts"]:
        add_item(d.get("folder_no"), "Dictionary", d["name"], d["dict_no"])

    for c in maps["counters"]:
        add_item(c.get("folder_no"), "Counter", c["name"], c["counter_no"])

    stamps_el = root.find("Stamps")
    for st in (stamps_el if stamps_el is not None else []):
        fn   = get_int(st, "Folder") or get_int(st, "FolderNo") or None
        sno  = get_text(st, "StampNo")
        snam = get_text(st, "Name") or f"Stamp_{sno}"
        add_item(fn, "Stamp", snam, sno)

    for dt in maps["data_types"]:
        add_item(dt.get("folder_no"), "Data Type", dt["name"], dt["dt_no"])

    # Render ───────────────────────────────────────────────────────────────
    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Folder Structure"), 15, 16, 5, bottom_border=True)
    total_objects = sum(len(v) for v in folder_items.values())
    doc.add_paragraph(
        f"{len(fn_nodes)} folders, {total_objects} objects. "
        "Each object is listed under its folder."
    )

    def render_folder(fno, depth):
        fname, _ = fn_nodes[fno]
        indent   = Cm(depth * 0.75)

        # Folder heading
        fp = doc.add_paragraph(style="List Bullet")
        fp.paragraph_format.left_indent = indent
        run = fp.add_run(fname)
        run.bold  = True
        run.font.color.rgb = _theme.docx_rgb("primary")
        fp.add_run(f"  (No: {show_id(fno)})")

        # Items inside this folder, sorted by type then name
        items = sorted(
            folder_items.get(fno, []),
            key=lambda x: (TYPE_ORDER.index(x[0]) if x[0] in TYPE_ORDER else 99, x[1].lower()),
        )
        for type_label, name, obj_no in items:
            ip = doc.add_paragraph(style="List Bullet")
            ip.paragraph_format.left_indent = Cm((depth + 1) * 0.75)
            ip.add_run(name).bold = True
            ip.add_run(f"  [{type_label}  No: {show_id(obj_no)}]"
                       ).font.color.rgb = _theme.docx_rgb("muted")

        # Recurse into child folders
        children = sorted(fn_childs.get(fno, []),
                          key=lambda c: fn_nodes[c][0].lower() if c in fn_nodes else "")
        for child_fno in children:
            render_folder(child_fno, depth + 1)

    # Start from root folders (parent = 0 / None)
    root_folders = sorted(fn_childs.get(0, []),
                          key=lambda c: fn_nodes[c][0].lower() if c in fn_nodes else "")
    for fno in root_folders:
        render_folder(fno, 0)

    # Objects with no folder (folder_no = None / 0)
    unfoldered = folder_items.get(0, []) + folder_items.get(None, [])
    if unfoldered:
        p = doc.add_paragraph()
        p.add_run("(No folder)").bold = True
        for type_label, name, obj_no in sorted(unfoldered, key=lambda x: x[1].lower()):
            ip = doc.add_paragraph(style="List Bullet")
            ip.paragraph_format.left_indent = Cm(0.75)
            ip.add_run(name).bold = True
            ip.add_run(f"  [{type_label}  No: {show_id(obj_no)}]"
                       ).font.color.rgb = _theme.docx_rgb("muted")


# ---------------------------------------------------------------------------
# Section: keyword dictionaries
# ---------------------------------------------------------------------------
def build_keyword_dicts(doc, maps, categories, section_no):
    kw_dicts = maps["kw_dicts"]
    if not kw_dicts:
        return

    # Build usage map: dict_no → [(cat_name, field_caption, is_multi)]
    td = maps["type_to_dict"]
    usage = {}
    for ctgry_no, cat_name, _folder_no, cat_elem in categories:
        fe = cat_elem.find("Fields")
        if fe is None:
            continue
        for field in fe.findall("Field"):
            try:
                tno = int(get_text(field, "TypeNo") or 0)
            except ValueError:
                continue
            if tno < 0 and tno in td:
                d = td[tno]
                cap = get_caption(field)
                usage.setdefault(d["dict_no"], []).append(
                    (cat_name, cap, d["is_multi"])
                )

    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Keyword Dictionaries"), 15, 16, 5, bottom_border=True)
    doc.add_paragraph(
        f"{len(kw_dicts)} keyword dictionaries. Each is a pick-list used by one "
        "or more category fields."
    )

    for d in sorted(kw_dicts, key=lambda x: x["name"]):
        blue_heading(doc, d["name"], level=2,
                     meta=f"(Dictionary No: {show_id(d['dict_no'])})")

        usages = usage.get(d["dict_no"], [])
        if usages:
            p = doc.add_paragraph()
            p.add_run("Used in: ").italic = True
            by_cat = {}
            for cat_name, field_cap, is_multi in usages:
                tag = f"{field_cap}{' (multi)' if is_multi else ''}"
                by_cat.setdefault(cat_name, []).append(tag)
            parts = [f"{cat}: {', '.join(fields)}" for cat, fields in by_cat.items()]
            p.add_run(";  ".join(parts))

        if d["keywords"]:
            t = doc.add_table(rows=1, cols=2)
            t.style = "Table Grid"
            hdr_cell(t.rows[0].cells, 0, "No")
            hdr_cell(t.rows[0].cells, 1, "Value")
            for idx, (kw_no, kw_val) in enumerate(
                sorted(d["keywords"], key=lambda x: x[0])
            ):
                fill = _theme.hex("grey") if idx % 2 == 1 else _theme.hex("white")
                rc = t.add_row().cells
                body_cell(rc, 0, str(kw_no), fill)
                body_cell(rc, 1, kw_val,     fill)
            set_col_widths(t, [1.8, 15.0])
        else:
            doc.add_paragraph().add_run("No keywords defined.").italic = True

        doc.add_paragraph()


# ---------------------------------------------------------------------------
# Document style setup
# ---------------------------------------------------------------------------
def setup_document_styles(doc):
    """Apply consistent typography across Normal, headings, and list styles."""
    from docx.oxml import OxmlElement

    # Normal / body text
    normal = doc.styles["Normal"]
    normal.font.name = _theme.font("body")
    normal.font.size = Pt(10)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after  = Pt(3)

    # Heading hierarchy
    for level, size, space_before, space_after in [
        (1, 15, 16, 5),
        (2, 12, 12, 3),
        (3, 10, 8,  2),
    ]:
        style = doc.styles[f"Heading {level}"]
        style.font.name  = _theme.font("body")
        style.font.size  = Pt(size)
        style.font.bold  = True
        style.font.color.rgb = _theme.docx_rgb("primary")
        style.paragraph_format.space_before  = Pt(space_before)
        style.paragraph_format.space_after   = Pt(space_after)
        style.paragraph_format.keep_with_next = True

    # H1 gets a bottom border rule
    h1_pPr = doc.styles["Heading 1"]._element.get_or_add_pPr()
    pBdr   = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), _theme.hex("h1_border"))
    pBdr.append(bottom)
    h1_pPr.append(pBdr)

    # List Bullet
    try:
        lb = doc.styles["List Bullet"]
        lb.font.name = _theme.font("body")
        lb.font.size = Pt(10)
    except KeyError:
        pass

    # Title (level 0)
    try:
        title_s = doc.styles["Title"]
        title_s.font.name  = _theme.font("body")
        title_s.font.size  = Pt(24)
        title_s.font.bold  = True
        title_s.font.color.rgb = _theme.docx_rgb("primary")
    except KeyError:
        pass


def add_footer(doc, doc_title="Therefore Documentation"):
    """Add a footer with the document title — no field codes to avoid Word prompts."""
    from docx.oxml import OxmlElement

    for section in doc.sections:
        footer = section.footer
        for p in footer.paragraphs:
            p.clear()
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.paragraph_format.space_before = Pt(0)
        fp.paragraph_format.space_after  = Pt(0)

        # Top border
        pPr  = fp._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        top  = OxmlElement("w:top")
        top.set(qn("w:val"),   "single")
        top.set(qn("w:sz"),    "4")
        top.set(qn("w:space"), "1")
        top.set(qn("w:color"), _theme.hex("footer_border"))
        pBdr.append(top)
        pPr.append(pBdr)

        r = fp.add_run(f"Therefore Documentation  —  {doc_title}")
        r.font.name      = _theme.font("body")
        r.font.size      = Pt(8)
        r.font.italic    = True
        r.font.color.rgb = _theme.docx_rgb("footer_text")


def add_toc(doc):
    """Insert a Table of Contents field (populated by Word on open)."""
    from docx.oxml import OxmlElement

    _manual_heading(doc, "Contents", 15, 16, 5, bottom_border=True)

    para = doc.add_paragraph()
    run  = para.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:r")
    placeholder_t = OxmlElement("w:t")
    placeholder_t.text = "(Open document and update fields to generate contents)"
    placeholder.append(placeholder_t)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, sep])
    para._p.append(placeholder)
    run2 = para.add_run()
    run2._r.append(end)

    run.font.name  = _theme.font("body")
    run.font.size  = Pt(10)
    doc.add_page_break()




# ---------------------------------------------------------------------------
# Section: category relationships
# ---------------------------------------------------------------------------
def _render_mermaid(mermaid_src, img_path, render_dir, width=1800, height=5000):
    """Render a mermaid string to PNG. Returns img_path or None."""
    try:
        from render_workflow import _find_mmdc, _autocrop
        import subprocess, tempfile, json as _json
        mmdc = _find_mmdc()
        if not mmdc:
            return None
        with tempfile.NamedTemporaryFile(mode="w", suffix=".mmd",
                                         delete=False, encoding="utf-8") as f:
            f.write(mermaid_src); tmp = f.name
        _puppeteer_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "puppeteer.json")
        cmd = mmdc.split() + [
            "-i", tmp, "-o", img_path,
            "--width", str(width), "--height", str(height),
            "--scale", "2", "--backgroundColor", "white",
        ]
        if os.path.exists(_puppeteer_cfg):
            try:
                _cfg = _json.loads(open(_puppeteer_cfg).read())
                if not _cfg.get("executablePath") or os.path.exists(_cfg["executablePath"]):
                    cmd += ["--puppeteerConfigFile", _puppeteer_cfg]
            except Exception:
                pass
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        os.unlink(tmp)
        if result.returncode == 0 and os.path.exists(img_path):
            _autocrop(img_path, padding=20)
            return img_path
    except Exception as e:
        print(f"Warning: diagram render failed ({e})")
    return None


def _embed_img(doc, img_path, max_w_cm=17.0, max_h_cm=20.0):
    """Embed a PNG in the document, constrained to max dimensions."""
    from PIL import Image as _PilImg
    _iw, _ih = _PilImg.open(img_path).size
    max_w, max_h = Cm(max_w_cm), Cm(max_h_cm)
    pic_kwargs = {"height": max_h} if _iw > 0 and (_ih / _iw) * max_w > max_h \
                 else {"width": max_w}
    pp = doc.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pp.add_run().add_picture(img_path, **pic_kwargs)
    add_render_note(doc)


def build_relationship_diagram(doc, root, maps, categories, section_no,
                                render_dir=None):
    import re as _re

    # ── Build lookup maps ──────────────────────────────────────────────────
    def _to_int(v):
        try: return int(v)
        except (TypeError, ValueError): return v

    cat_no_to_name = {
        _to_int(cno): name
        for cno, name, _fno, _el in categories
    }

    # data type no → target category no  (TheCat-N format)
    dt_to_cat_no = {}
    _dt_el = root.find("Datatypes") or root.find("DataTypes")
    for dt in (_dt_el if _dt_el is not None else []):
        tno = get_text(dt, "TypeNo")
        tbl = get_text(dt, "TableName") or ""
        m = _re.match(r"TheCat(-?\d+)", tbl)
        if m and tno:
            dt_to_cat_no[tno] = m.group(1)

    # indexing profile no → (target category no, profile name)
    ixp_to_cat = {}
    ixp_name = {}
    for p in (root.find("IxProfiles") or []):
        pno   = get_text(p, "IndexingProfileNo")
        def_el = p.find("DefId")
        tgt_no = get_text(def_el, "DefNo") if def_el is not None else None
        pname = get_text(p, "Name") or f"Profile_{pno}"
        if pno and tgt_no:
            ixp_to_cat[pno] = tgt_no
            ixp_name[pno] = pname

    # ── Collect edges ──────────────────────────────────────────────────────
    # (src_id, tgt_id, label, edge_type)
    # edge_type: "cat_cat" | "cat_kw" | "eform_cat"
    edges = []
    seen_edges = set()

    def add_edge(src, tgt, label, etype):
        key = (src, tgt, etype)
        if key not in seen_edges:
            seen_edges.add(key)
            edges.append((src, tgt, label, etype))

    kw_dicts = {d["dict_no"]: d["name"] for d in maps["kw_dicts"]}
    td = maps["type_to_dict"]

    for cno, cname, _fno, cat_elem in categories:
        fields_el = cat_elem.find("Fields")
        for f in (fields_el if fields_el is not None else []):
            tno_str = get_text(f, "TypeNo")
            if not tno_str:
                continue
            try:
                tno = int(tno_str)
            except ValueError:
                continue

            # Category → Keyword Dictionary  (td keys are integers)
            if tno in td:
                d = td[tno]
                add_edge(f"cat_{_to_int(cno)}", f"kw_{d['dict_no']}",
                         "", "cat_kw")

            # Category → Category (via referenced data type)
            if tno_str in dt_to_cat_no:
                target_cno_str = dt_to_cat_no[tno_str]
                target_cno_int = _to_int(target_cno_str)
                if target_cno_int != _to_int(cno) and target_cno_int in cat_no_to_name:
                    field_col = get_text(f, "ColName") or ""
                    add_edge(f"cat_{_to_int(cno)}", f"cat_{target_cno_int}",
                             field_col, "cat_cat")

    # eForm → Category
    ef_el = root.find("EForms")
    ef_names = {}
    for ef in (ef_el if ef_el is not None else []):
        ef_name_str = ef.findtext("FName") or "EForm"
        ef_no       = ef.findtext("FNo") or "x"
        ef_names[ef_no] = ef_name_str
        pno         = ef.findtext("IxProfNo")
        cat_no_str  = ixp_to_cat.get(pno)
        if cat_no_str:
            cat_no_int = _to_int(cat_no_str)
            if cat_no_int in cat_no_to_name:
                prof_name = ixp_name.get(pno, "")
                add_edge(f"ef_{ef_no}", f"cat_{cat_no_int}", prof_name, "eform_cat")

    cat_cat_edges  = [(s, t, l) for s, t, l, e in edges if e == "cat_cat"]
    cat_kw_edges   = [(s, t, l) for s, t, l, e in edges if e == "cat_kw"]
    eform_edges    = [(s, t, l) for s, t, l, e in edges if e == "eform_cat"]

    if not edges:
        return

    def safe_id(s):
        return _re.sub(r"[^a-zA-Z0-9_]", "_", str(s))
    def mml(t):
        return str(t).replace('"', "'")
    def cat_name(node_id):
        return cat_no_to_name.get(_to_int(node_id[4:]), node_id)

    # ── Section heading ───────────────────────────────────────────────────
    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Category Relationships"), 15, 16, 5, bottom_border=True)

    # ── Unified ecosystem diagram ─────────────────────────────────────────
    # Build a single Mermaid diagram with categories, keyword dicts, eForms
    lines = ["%%{init: {'flowchart': {'nodeSpacing': 45, 'rankSpacing': 70}}}%%",
             "flowchart LR"]
    declared = set()

    def declare_node(nid, label, shape):
        if nid in declared:
            return
        declared.add(nid)
        sid = safe_id(nid)
        ml = mml(label)
        if shape == "diamond":
            lines.append(f'    {sid}{{"{ml}"}}')
        elif shape == "hexagon":
            lines.append(f'    {sid}{{"{{"{ml}"}}}}')
        elif shape == "parallelogram":
            lines.append(f'    {sid}[/{ml}/]')
        else:
            lines.append(f'    {sid}["{ml}"]')

    # Category → Category edges (solid)
    for src, tgt, lbl in cat_cat_edges:
        declare_node(src, cat_name(src), "round")
        declare_node(tgt, cat_name(tgt), "round")
    for src, tgt, lbl in cat_cat_edges:
        sid, tid = safe_id(src), safe_id(tgt)
        if lbl:
            lines.append(f'    {sid} -->|"{mml(lbl)}"| {tid}')
        else:
            lines.append(f'    {sid} --> {tid}')

    # Category → Keyword Dictionary edges (dotted)
    for src, tgt, lbl in cat_kw_edges:
        declare_node(src, cat_name(src), "round")
        kw_no = int(tgt[3:])
        kw_label = kw_dicts.get(kw_no, f"Dict {kw_no}")
        declare_node(tgt, kw_label, "diamond")
    for src, tgt, lbl in cat_kw_edges:
        sid, tid = safe_id(src), safe_id(tgt)
        lines.append(f'    {sid} -.-> {tid}')

    # eForm → Category edges (dashed)
    for src, tgt, lbl in eform_edges:
        ef_no = src[3:]
        ef_label = ef_names.get(ef_no, f"EForm {ef_no}")
        declare_node(src, ef_label, "hexagon")
        declare_node(tgt, cat_name(tgt), "round")
    for src, tgt, lbl in eform_edges:
        sid, tid = safe_id(src), safe_id(tgt)
        if lbl:
            lines.append(f'    {sid} -.->|"{mml(lbl)}"| {tid}')
        else:
            lines.append(f'    {sid} -.-> {tid}')

    # Class definitions from theme
    mm = _theme.mermaid
    def _mm_cls(key, default_fill, default_stroke):
        cfg = mm.get(key, {})
        fill = cfg.get("fill", default_fill)
        stroke = cfg.get("stroke", default_stroke)
        return f"    classDef {key} fill:#{fill},stroke:#{stroke},stroke-width:2px"

    lines.append(_mm_cls("catNode", "D6E4F0", "1F4E79"))
    lines.append(_mm_cls("kwNode", "E8D6F0", "6C3483"))
    lines.append(_mm_cls("efNode", "D6F0E8", "1E8449"))

    # Assign classes
    for src, tgt, _ in cat_cat_edges:
        lines.append(f"    class {safe_id(src)} catNode")
        lines.append(f"    class {safe_id(tgt)} catNode")
    for src, tgt, _ in cat_kw_edges:
        lines.append(f"    class {safe_id(src)} catNode")
        lines.append(f"    class {safe_id(tgt)} kwNode")
    for src, tgt, _ in eform_edges:
        lines.append(f"    class {safe_id(src)} efNode")
        lines.append(f"    class {safe_id(tgt)} catNode")

    mermaid_src = "\n".join(lines)

    if render_dir and declared:
        img = _render_mermaid(mermaid_src,
                              os.path.join(render_dir, "ecosystem_diagram.png"),
                              render_dir, width=1600, height=4000)
        if img:
            _embed_img(doc, img)

    # ── 1. Category cross-references ──────────────────────────────────────
    if cat_cat_edges:
        blue_heading(doc, "Category Cross-references", level=2)
        doc.add_paragraph(
            "Categories that reference fields from another category via a linked data type."
        )

        t = doc.add_table(rows=1, cols=3)
        t.style = "Table Grid"
        for i, h in enumerate(["Source Category", "Via Field", "Target Category"]):
            hdr_cell(t.rows[0].cells, i, h)
        from itertools import groupby
        sorted_edges = sorted(cat_cat_edges, key=lambda e: cat_name(e[0]))
        for idx, (src, tgt, lbl) in enumerate(sorted_edges):
            fill = _theme.hex("grey") if idx % 2 else _theme.hex("white")
            rc = t.add_row().cells
            body_cell(rc, 0, cat_name(src), fill, bold=True)
            body_cell(rc, 1, lbl or "—", fill, mono=True, pt=8)
            body_cell(rc, 2, cat_name(tgt), fill, bold=True)
        set_col_widths(t, [6.0, 4.5, 6.0])
        doc.add_paragraph()

    # ── 2. Keyword dictionary usage ───────────────────────────────────────
    if cat_kw_edges:
        blue_heading(doc, "Keyword Dictionary Usage", level=2)
        doc.add_paragraph(
            "Categories that use a keyword (pick-list) dictionary for one or more fields."
        )
        from collections import defaultdict
        kw_to_cats = defaultdict(set)
        for src, tgt, _ in cat_kw_edges:
            kw_no = int(tgt[3:])
            kw_to_cats[kw_no].add(_to_int(src[4:]))

        t = doc.add_table(rows=1, cols=2)
        t.style = "Table Grid"
        for i, h in enumerate(["Keyword Dictionary", "Used by Categories"]):
            hdr_cell(t.rows[0].cells, i, h)
        for idx, (kw_no, cat_nos) in enumerate(
                sorted(kw_to_cats.items(), key=lambda x: kw_dicts.get(x[0], ""))):
            fill = _theme.hex("grey") if idx % 2 else _theme.hex("white")
            rc = t.add_row().cells
            body_cell(rc, 0, kw_dicts.get(kw_no, f"Dict {kw_no}"), fill, bold=True)
            cat_list = ", ".join(sorted(
                cat_no_to_name.get(cn, str(cn)) for cn in cat_nos
            ))
            body_cell(rc, 1, cat_list, fill)
        set_col_widths(t, [5.5, 11.0])
        doc.add_paragraph()

    # ── 3. eForm submissions ──────────────────────────────────────────────
    if eform_edges:
        blue_heading(doc, "eForm Submissions", level=2)
        doc.add_paragraph(
            "eForms that submit documents into a Therefore category via an indexing profile."
        )
        t = doc.add_table(rows=1, cols=2)
        t.style = "Table Grid"
        for i, h in enumerate(["eForm", "Submits to Category"]):
            hdr_cell(t.rows[0].cells, i, h)
        for idx, (src, tgt, lbl) in enumerate(eform_edges):
            fill = _theme.hex("grey") if idx % 2 else _theme.hex("white")
            rc   = t.add_row().cells
            ef_no   = src[3:]
            ef_name = ef_names.get(ef_no, f"EForm {ef_no}")
            body_cell(rc, 0, ef_name, fill, bold=True)
            body_cell(rc, 1, cat_name(tgt), fill)
        set_col_widths(t, [8.0, 8.5])
        doc.add_paragraph()



# ---------------------------------------------------------------------------
# Generic list section
# ---------------------------------------------------------------------------
REPORT_TYPE_LABELS = {
    "1": "Category",
    "2": "Workflow",
    "3": "System",
    "4": "Custom",
}


def build_reports_section(doc, items, section_no):
    if not items:
        return
    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Report Definitions"), 15, 16, 5,
                    bottom_border=True)
    doc.add_paragraph(f"Total report definitions: {len(items)}")

    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    for i, h in enumerate(["No", "Name", "Type", "Path"]):
        hdr_cell(t.rows[0].cells, i, h)

    for idx, r in enumerate(sorted(items, key=lambda x: get_name(x) or "")):
        fill    = _theme.hex("grey") if idx % 2 else _theme.hex("white")
        rno     = get_text(r, "RptDefNo") or get_text(r, "ReportNo") or ""
        name    = get_text(r, "RptName") or get_name(r) or ""
        rtype   = REPORT_TYPE_LABELS.get(get_text(r, "RptType") or "", "—")
        path    = get_text(r, "RptPath") or "—"
        rc      = t.add_row().cells
        body_cell(rc, 0, show_id(rno), fill)
        body_cell(rc, 1, name, fill, bold=True)
        body_cell(rc, 2, rtype, fill)
        body_cell(rc, 3, path, fill, pt=8)

    set_col_widths(t, [1.2, 6.0, 3.0, 6.2])


def build_list_section(doc, items, section_no, title, id_field):
    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, title), 15, 16, 5, bottom_border=True)
    doc.add_paragraph(f"Total: {len(items)}")
    for item in items:
        item_no = get_text(item, id_field)
        name    = get_name(item) or f"{title}_{item_no}"
        lp      = doc.add_paragraph(style="List Bullet")
        lp.add_run(name).bold = True
        lp.add_run(f"  ({id_field}: {show_id(item_no)})")


# ---------------------------------------------------------------------------
# Section: script inventory
# ---------------------------------------------------------------------------
def build_script_inventory(doc, profiles, root, section_no, maps=None,
                           ai_url=None, ai_model=None, ai_key="lm-studio"):
    """Appendix collecting all custom scripts from profiles, eForms, and workflows."""

    entries = []

    # ── A. Indexing profiles ───────────────────────────────────────────────
    for p in profiles:
        lang = SCRIPT_LANG_LABELS.get(p["lang"], "") if p["lang"] else ""
        source = f"Profile: {p['name']}"
        if p["target_cat"]:
            source += f" → {p['target_cat']}"

        if p["init_script"]:
            entries.append({
                "source": source,
                "context": "Init Script",
                "lang": lang,
                "code": p["init_script"],
            })

        for a in p.get("assignments", []):
            if a["expr"]:
                entries.append({
                    "source": source,
                    "context": f"Field: {a['field']} — Expression",
                    "lang": lang,
                    "code": a["expr"],
                })
            if a["check"]:
                entries.append({
                    "source": source,
                    "context": f"Field: {a['field']} — Check Condition",
                    "lang": lang,
                    "code": a["check"],
                })

    # ── B. eForms ──────────────────────────────────────────────────────────
    ef_parent = root.find("EForms")
    for ef in (ef_parent if ef_parent is not None else []):
        ef_name = ef.findtext("FName") or "EForm"
        fdef_text = ef.findtext("FDef")
        if not fdef_text:
            continue
        try:
            fdef = json.loads(fdef_text)
        except Exception:
            continue

        def _walk_ef(components, page="", datagrid=""):
            for c in components:
                t = c.get("type", "")
                label = c.get("label", "") or c.get("title", "")
                key = c.get("key", "")
                ctx = f"{ef_name}"
                if page:
                    ctx += f" / {page}"
                if datagrid:
                    ctx += f" / {datagrid}"
                ctx += f" — {label or key}"

                if t == "panel":
                    _walk_ef(c.get("components") or [], page=label or page, datagrid=datagrid)
                    continue
                elif t == "columns":
                    for col in (c.get("columns") or []):
                        _walk_ef(col.get("components") or [], page=page, datagrid=datagrid)
                    continue
                elif t == "datagrid":
                    _walk_ef(c.get("components") or [], page=page, datagrid=label)
                    continue

                cc = (c.get("customConditional") or "").strip()
                if cc:
                    entries.append({"source": ctx, "context": "Custom Conditional", "lang": "JavaScript", "code": cc})

                calc = (c.get("calculateValue") or "").strip()
                if calc:
                    entries.append({"source": ctx, "context": "Calculated Value", "lang": "JavaScript", "code": calc})

                val = c.get("validate") or {}
                vcustom = (val.get("custom") or "").strip()
                if vcustom:
                    entries.append({"source": ctx, "context": "Custom Validation", "lang": "JavaScript", "code": vcustom})

                for l in (c.get("logic") or []):
                    tr = l.get("trigger") or {}
                    tr_type = tr.get("type", "")
                    tr_desc = tr.get("event", "") if tr_type == "event" else ""
                    actions = []
                    for a in (l.get("actions") or []):
                        atype = a.get("type", "")
                        prop = (a.get("property") or {}).get("value", "")
                        val = str(a.get("value", "") or a.get("customAction", "")).strip()
                        if prop:
                            actions.append(f"set {prop} = {val}")
                        else:
                            actions.append(f"{atype}: {val}")
                    if actions:
                        code = f"Trigger: {tr_desc or tr_type}\n" + "\n".join(actions)
                        entries.append({"source": ctx, "context": "Logic", "lang": "JavaScript", "code": code})

        _walk_ef(fdef.get("components") or [])

    # ── C. Workflows ───────────────────────────────────────────────────────
    wf_parent = root.find("WFProcesses") or root.find("Workflows")
    if wf_parent is not None:
        for wf in (wf_parent.findall("WFProcess") or wf_parent.findall("Workflow")):
            wf_name = get_name(wf) or "Workflow"
            tasks_el = wf.find("Tasks")
            if tasks_el is None:
                continue
            for task in tasks_el.findall("T") or tasks_el.findall("Task"):
                tname = get_name(task) or f"Task {get_text(task, 'TaskNo')}"
                # Task-level script (often found in automatic tasks)
                script = (task.findtext("Script") or "").strip()
                if script:
                    entries.append({
                        "source": f"{wf_name} / {tname}",
                        "context": "Task Script",
                        "lang": "",
                        "code": script,
                    })
                # Transition conditions
                trans_el = task.find("Transitions")
                if trans_el is not None:
                    for tr in trans_el.findall("TR"):
                        cond = (get_text(tr, "Condition") or "").strip()
                        if cond and len(cond) > 3 and not cond.replace(" ", "").isdigit():
                            entries.append({
                                "source": f"{wf_name} / {tname}",
                                "context": "Transition Condition",
                                "lang": "",
                                "code": cond,
                            })

    if not entries:
        return

    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Script Inventory"), 15, 16, 5,
                    bottom_border=True)
    doc.add_paragraph(
        f"{len(entries)} custom script(s), calculated fields, and conditions "
        "found across indexing profiles, eForms, and workflows."
    )

    fno_map = (maps or {}).get("field_no_map", {})

    # Pre-count scripts that will need AI summaries so we can show N of M progress
    if ai_url:
        _ai_candidates = [
            e for e in entries
            if len([l for l in resolve_field_refs(e["code"], fno_map).split("\n") if l.strip()]) > 4
        ]
        if _ai_candidates:
            print(f"Generating AI summaries for {len(_ai_candidates)} script(s) ...")
    else:
        _ai_candidates = []
    _ai_done = 0

    for e in entries:
        meta = f"{e['context']} ({e['lang']})" if e['lang'] else e['context']
        blue_heading(doc, e["source"], level=3, meta=meta)

        resolved_code = resolve_field_refs(e["code"], fno_map)
        non_empty_lines = [l for l in resolved_code.split("\n") if l.strip()]

        rows = [
            ("Context",  e["context"]),
            ("Language", e["lang"] or "—"),
        ]

        if ai_url and len(non_empty_lines) > 4:
            _ai_done += 1
            print(f"  AI summary {_ai_done}/{len(_ai_candidates)}: {e['source'][:60]}")
            try:
                from ai_summary import summarize_script
                summary = summarize_script(
                    resolved_code,
                    context=f"{e['source']} — {e['context']}",
                    ai_url=ai_url,
                    ai_model=ai_model,
                    api_key=ai_key,
                )
                rows.append(("Summary", summary))
            except Exception as exc:
                print(f"  Script summary failed: {exc}")

        kv_table(doc, rows, lw=3.5, vw=13.0)
        add_code_block(doc, resolved_code)


# ---------------------------------------------------------------------------
# Section: server configuration
# ---------------------------------------------------------------------------
def _kv_table(doc, rows, lw=5.0, vw=11.5):
    """Render a two-column label/value table."""
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = "Table Grid"
    for i, (label, value) in enumerate(rows):
        fill = _theme.hex("grey") if i % 2 else _theme.hex("white")
        body_cell(t.rows[i].cells, 0, label, fill, bold=True)
        body_cell(t.rows[i].cells, 1, value or "—", fill)
    set_col_widths(t, [lw, vw])


def _sub_heading(doc, text):
    p = doc.add_paragraph()
    p.add_run(text).bold = True
    p.runs[0].font.size = Pt(10)
    p.runs[0].font.color.rgb = _theme.docx_rgb("primary")
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(3)


def build_server_info_section(doc, server_info, xml_root, categories, section_no):
    """
    Render the Server Configuration section.

    API-derived details (server_info dict) are shown when provided.
    Retention policies are always rendered if present in the XML.
    """
    ret_el = xml_root.find("RetentionPolicies")
    policies = list(ret_el) if ret_el is not None else []

    # Nothing to show
    if not server_info and not policies:
        return

    doc.add_page_break()
    _manual_heading(doc, sec_heading(section_no, "Server Configuration"),
                    15, 16, 5, bottom_border=True)

    si = server_info or {}

    if si:
        # --- System ---
        _sub_heading(doc, "System")
        sys_rows = []
        if si.get("api_url"):            sys_rows.append(("API URL",           si["api_url"]))
        if si.get("api_tenant"):         sys_rows.append(("Tenant",            si["api_tenant"]))
        if si.get("server_name"):        sys_rows.append(("Server Name",       si["server_name"]))
        if si.get("service_version"):    sys_rows.append(("Service Version",   si["service_version"]))
        if si.get("customer_id"):        sys_rows.append(("Customer ID",       si["customer_id"]))
        if si.get("tenant_name") and si.get("tenant_name") != si.get("api_tenant"):
            sys_rows.append(("Tenant Name",       si["tenant_name"]))
        if si.get("region"):             sys_rows.append(("Region",            si["region"]))
        if si.get("locale"):             sys_rows.append(("Locale",            si["locale"]))
        if si.get("admin_email"):        sys_rows.append(("Admin Email",       si["admin_email"]))
        if si.get("domain_names"):       sys_rows.append(("Domain Names",      ", ".join(si["domain_names"])))
        if si.get("default_domain"):     sys_rows.append(("Default Domain",    si["default_domain"]))
        if si.get("terms_url"):          sys_rows.append(("Terms of Service",  si["terms_url"]))
        if sys_rows:
            _kv_table(doc, sys_rows)

        # --- Database ---
        db_rows = []
        if si.get("db_connection"): db_rows.append(("Server / Database",   si["db_connection"]))
        if si.get("db_username"):   db_rows.append(("Username",            si["db_username"]))
        if si.get("db_options"):    db_rows.append(("Connection Options",  si["db_options"]))
        if db_rows:
            _sub_heading(doc, "Database")
            _kv_table(doc, db_rows)

        # --- SMTP ---
        smtp_rows = []
        if si.get("smtp_server"):    smtp_rows.append(("Server",      si["smtp_server"]))
        if si.get("smtp_sender"):    smtp_rows.append(("Sender",      si["smtp_sender"]))
        if "smtp_ssl" in si:         smtp_rows.append(("SSL",         "Yes" if si["smtp_ssl"] else "No"))
        if si.get("smtp_auth_user"): smtp_rows.append(("Auth User",   si["smtp_auth_user"]))
        if smtp_rows:
            _sub_heading(doc, "Email (SMTP)")
            _kv_table(doc, smtp_rows)

        # --- Storage ---
        storage_rows = []
        _storage_labels = [
            ("buffer_path",   "Buffer"),
            ("fulltext_path", "Full-Text Index"),
            ("preview_blob",  "Preview Storage"),
            ("fulltext_blob", "Full-Text Blob"),
            ("webviewer_blob","Web Viewer"),
            ("storage_blob",  "Document Storage Blob"),
            ("cache_path",    "Cache"),
            ("report_server", "Report Server"),
        ]
        for key, label in _storage_labels:
            if si.get(key):
                storage_rows.append((label, si[key]))
        if storage_rows:
            _sub_heading(doc, "Storage Areas")
            t = doc.add_table(rows=1, cols=2)
            t.style = "Table Grid"
            hdr_cell(t.rows[0].cells, 0, "Area")
            hdr_cell(t.rows[0].cells, 1, "Path / URL")
            for idx, (label, path) in enumerate(storage_rows):
                fill = _theme.hex("grey") if idx % 2 else _theme.hex("white")
                rc = t.add_row().cells
                body_cell(rc, 0, label, fill, bold=True)
                body_cell(rc, 1, path,  fill, pt=8)
            set_col_widths(t, [4.0, 12.5])

        # --- Licensed storage ---
        lic_rows = []
        if si.get("storage_licensed_gb"):
            lic_rows.append(("Licensed Storage", f"{si['storage_licensed_gb']} GB"))
        if si.get("storage_used_gb"):
            lic_rows.append(("Used / Threshold", f"{si['storage_used_gb']} GB"))
        if lic_rows:
            _sub_heading(doc, "Licensed Storage")
            _kv_table(doc, lic_rows)

    # --- Retention Policies (from XML) ---
    if policies:
        cat_name_map = {str(no): name for no, name, _, _ in categories}
        _sub_heading(doc, "Retention Policies")
        t = doc.add_table(rows=1, cols=5)
        t.style = "Table Grid"
        for i, h in enumerate(["Name", "Duration", "Starts From", "On Expiry", "Applied To"]):
            hdr_cell(t.rows[0].cells, i, h)

        for idx, pol in enumerate(policies):
            fill  = _theme.hex("grey") if idx % 2 else _theme.hex("white")
            name  = pol.findtext("Name") or "—"
            months = pol.findtext("Months") or ""
            try:
                duration = f"{int(months)} months ({int(months)/12:.4g} years)"
            except (ValueError, TypeError):
                duration = f"{months} months" if months else "—"
            start  = pol.findtext("Starting") or "—"
            purge  = pol.findtext("Purge") == "1"
            delete = pol.findtext("DeleteDisk") == "1"
            if purge and delete:
                expiry = "Purge index + delete files"
            elif purge:
                expiry = "Purge index record"
            elif delete:
                expiry = "Delete files from disk"
            else:
                expiry = "No action"

            sub_cats = pol.findall("SubCtgrys/SubCtgry")
            active = [s for s in sub_cats if s.findtext("NoRetention") != "1"]
            if active:
                cat_names = [cat_name_map.get(s.findtext("CtgryNo") or "", s.findtext("CtgryNo") or "?")
                             for s in active]
                applied = f"{len(active)} categories: {', '.join(sorted(cat_names))}"
            else:
                applied = "None assigned"

            rc = t.add_row().cells
            body_cell(rc, 0, name,     fill, bold=True)
            body_cell(rc, 1, duration, fill)
            body_cell(rc, 2, start,    fill)
            body_cell(rc, 3, expiry,   fill)
            body_cell(rc, 4, applied,  fill, pt=8)

        set_col_widths(t, [3.5, 3.5, 3.0, 3.5, 3.0])


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def build_ai_summary_section(doc, summary_text: str, section_no):
    _manual_heading(doc, sec_heading(section_no, "System Summary"),
                    15, 16, 5, bottom_border=True)
    for para in summary_text.split("\n"):
        para = para.strip()
        if para:
            doc.add_paragraph(para)


ALL_SECTIONS = [
    "ai_summary",
    "server_info",
    "categories",
    "indexing_profiles",
    "workflows",
    "eforms",
    "relationships",
    "folder_structure",
    "keyword_dicts",
    "queries",
    "reports",
    "stamps",
    "script_inventory",
]


def generate(
    xml_path: str,
    output_path: str,
    *,
    skip_eforms: bool = False,
    no_images:   bool = False,
    body_only:   bool = False,
    render_dir:  str  = None,
    data_dir:    str  = None,
    sections:      list = None,
    start_section: int  = 1,
    log_fn=None,
    theme: Theme = None,
    img_format: str = "png",
    server_info: dict = None,
    ai_url: str = None,
    ai_model: str = None,
    ai_key: str = "lm-studio",
) -> list:
    """
    Generate a Therefore documentation Word document from an XML export.

    Returns a list of warning strings (empty if none).  Raises on fatal errors.
    ``log_fn``, if supplied, is called with each progress message string instead
    of printing to stdout.
    ``sections`` is a list of section keys (see ALL_SECTIONS).  None means all.
    """
    import sys as _sys
    import types

    global _theme
    _theme = theme if theme is not None else DEFAULT_THEME

    active_sections = set(sections) if sections is not None else set(ALL_SECTIONS)
    # Respect legacy skip_eforms flag
    if skip_eforms:
        active_sections.discard("eforms")

    args = types.SimpleNamespace(
        xml              = xml_path,
        output           = output_path,
        render_dir       = render_dir,
        data_dir         = data_dir,
        no_images        = no_images,
        skip_eforms      = skip_eforms,
        body_only        = body_only,
        active_sections  = active_sections,
        start_section    = start_section,
        img_format       = img_format,
        server_info      = server_info,
        ai_url           = ai_url,
        ai_model         = ai_model,
        ai_key           = ai_key,
    )

    warnings = []

    if log_fn:
        class _Tee:
            def write(self, s):
                if s.strip():
                    log_fn(s.rstrip())
                _sys.__stdout__.write(s)
            def flush(self):
                _sys.__stdout__.flush()
            def isatty(self):
                return False
        old_stdout = _sys.stdout
        _sys.stdout = _Tee()
        try:
            _main_impl(args, warnings)
        finally:
            _sys.stdout = old_stdout
    else:
        _main_impl(args, warnings)

    return warnings


# ---------------------------------------------------------------------------
# Main implementation
# ---------------------------------------------------------------------------
def _main_impl(args, warnings=None):
    if warnings is None:
        warnings = []

    print(f"Parsing {args.xml} ...")
    tree = ET.parse(args.xml)
    root = tree.getroot()

    print("Building lookup maps ...")
    maps = parse_lookup_maps(root)

    categories = parse_categories(root)
    queries    = root.findall(".//QueryTemplate") or root.findall(".//Query")
    reports    = root.findall(".//ReportDefinition")
    stamps     = root.findall(".//Stamp")

    # Count workflows for overview (handle both element name variants)
    _wfp = root.find("WFProcesses")
    if _wfp is None:
        _wfp = root.find("Workflows")
    if _wfp is not None:
        workflows = _wfp.findall("WFProcess") or _wfp.findall("Workflow")
    else:
        workflows = []

    profiles = parse_ix_profiles(root, maps, include_eforms=not args.skip_eforms)

    # Optional folder-path enrichment from REST pull (overrides XML-derived paths)
    if args.data_dir:
        pf = Path(args.data_dir) / "target_paths.json"
        if pf.exists():
            rest_paths = json.loads(pf.read_text())
            print(f"Loaded REST folder paths for {len(rest_paths)} categories.")
            # These are already keyed by category number as string — used in overview only

    # Image rendering
    include_images = not args.no_images
    render_dir     = None
    _tmp_dir       = None
    pil_font       = None   # shared PIL font for category + workflow renders
    img_format     = getattr(args, "img_format", "png")

    if include_images:
        render_dir = args.render_dir or tempfile.mkdtemp(prefix="therefore_renders_")
        if args.render_dir:
            os.makedirs(render_dir, exist_ok=True)
        else:
            _tmp_dir = render_dir

        if img_format == "svg":
            # Generate SVGs as primary output, but also generate PNGs for Word embedding
            try:
                from render_categories_svg import render_simple_category as render_svg_simple
                from render_categories_svg import render_tabbed_category as render_svg_tabbed
                print(f"Rendering {len(categories)} categories ...")
                for _no, name, _fno, cat_elem in categories:
                    if cat_has_tabs(cat_elem):
                        render_svg_tabbed(name, cat_elem, render_dir, theme=_theme)
                    else:
                        render_svg_simple(name, cat_elem, render_dir, theme=_theme)
                print("SVG rendering complete.")
            except ImportError as e:
                msg = f"SVG rendering unavailable ({e}) — continuing without images."
                print(f"Warning: {msg}")
                warnings.append(msg)
                include_images = False

            # Also generate PNGs for Word embedding
            if include_images:
                try:
                    from render_categories import render_simple_category, render_tabbed_category
                    from PIL import ImageFont
                    try:
                        pil_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 11)
                    except Exception:
                        try:
                            pil_font = ImageFont.truetype("arial.ttf", 11)
                        except Exception:
                            pil_font = ImageFont.load_default()
                    print(f"Rendering PNG fallbacks for Word embedding ...")
                    for _no, name, _fno, cat_elem in categories:
                        if cat_has_tabs(cat_elem):
                            render_tabbed_category(name, cat_elem, render_dir, pil_font)
                        else:
                            render_simple_category(name, cat_elem, render_dir, pil_font)
                except ImportError:
                    pass  # PNG fallback non-critical
        else:
            try:
                from render_categories import render_simple_category, render_tabbed_category
                from PIL import ImageFont
                try:
                    pil_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 11)
                except Exception:
                    try:
                        pil_font = ImageFont.truetype("arial.ttf", 11)
                    except Exception:
                        pil_font = ImageFont.load_default()

                print(f"Rendering {len(categories)} categories ...")
                for _no, name, _fno, cat_elem in categories:
                    if cat_has_tabs(cat_elem):
                        render_tabbed_category(name, cat_elem, render_dir, pil_font)
                    else:
                        render_simple_category(name, cat_elem, render_dir, pil_font)
                print("Rendering complete.")

            except ImportError as e:
                msg = f"Image rendering unavailable ({e}) — continuing without images."
                print(f"Warning: {msg}")
                warnings.append(msg)
                include_images = False

    # Build document
    print("Building document ...")
    doc = Document()
    setup_document_styles(doc)

    # Narrow margins for A4 — more table space
    for section in doc.sections:
        section.left_margin   = Cm(1.8)
        section.right_margin  = Cm(1.8)
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    # Derive a short title from the XML filename for header/footer
    xml_stem = os.path.splitext(os.path.basename(args.xml))[0]
    doc_title = xml_stem.replace("TheConfiguration-", "").replace("TheConfiguration", "").strip(" -_") or "Therefore Documentation"

    if not args.body_only:
        add_footer(doc, doc_title)

        # Title page
        t = _manual_heading(doc, "Therefore Implementation Details", 22, 0, 8)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub = doc.add_paragraph(doc_title)
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub.runs[0].bold = True
        sub.runs[0].font.size = Pt(14)
        sub.runs[0].font.color.rgb = _theme.docx_rgb("primary")
        dp = doc.add_paragraph(f"Generated: {date.today().isoformat()}")
        dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dp.runs[0].font.color.rgb = _theme.docx_rgb("muted")
        doc.add_page_break()

        add_render_note(doc)

    _ef_el = root.find("EForms")
    _eforms_count = len(list(_ef_el)) if _ef_el is not None else 0

    sec    = getattr(args, "active_sections", set(ALL_SECTIONS))
    _start = getattr(args, "start_section", 1)

    # start_section=0 means "no section numbers" (used when inserting into a wrapper).
    # Any value >= 1 enables numbering from that section.
    if _start == 0:
        sn = None   # no numbers
    else:
        sn = _start - 1   # incremented before each use

    def _next_sn():
        """Return the next section number, or None if numbering is disabled."""
        nonlocal sn
        if sn is None:
            return None
        sn += 1
        return sn

    # AI Summary (generated before overview so section number comes first)
    _ai_summary_text = None
    if "ai_summary" in sec and getattr(args, "ai_url", None):
        try:
            from ai_summary import generate_ai_summary
            _eforms_list = [
                {"name": (get_name(ef) or ef.findtext("FName") or "")}
                for ef in (root.find("EForms") or [])
            ]
            _ai_summary_text = generate_ai_summary(
                categories, workflows, profiles, _eforms_list, maps,
                server_info    = getattr(args, "server_info", None),
                ai_url         = args.ai_url,
                ai_model       = getattr(args, "ai_model", None),
                api_key        = getattr(args, "ai_key", "lm-studio"),
                log_fn         = print,
            )
        except Exception as exc:
            warnings.append(f"AI summary failed: {exc}")
            print(f"AI summary failed: {exc}")

    # Overview
    build_overview(doc, [
        ("Categories",         len(categories)),
        ("Workflow Processes",  len(workflows)),
        ("eForms",             _eforms_count),
        ("Indexing Profiles",   len(profiles)),
        ("Keyword Dictionaries", len(maps["kw_dicts"])),
        ("Folders",             len(maps["folder_paths"])),
        ("Queries",             len(queries)),
        ("Report Definitions",  len(reports)),
        ("Stamps",              len(stamps)),
    ], section_no=_next_sn())
    doc.add_page_break()

    # AI Summary section
    if "ai_summary" in sec and _ai_summary_text:
        build_ai_summary_section(doc, _ai_summary_text, _next_sn())

    # Server Configuration
    if "server_info" in sec:
        _si = getattr(args, "server_info", None)
        build_server_info_section(doc, _si, root, categories, _next_sn())

    # Categories
    if "categories" in sec:
        build_categories(doc, categories, render_dir, include_images, maps,
                         section_no=_next_sn(), img_format=img_format)

    # Indexing Profiles
    if "indexing_profiles" in sec and profiles:
        build_ix_profiles(doc, profiles, _next_sn(), maps=maps)

    # Workflows
    if "workflows" in sec:
        wf_parent = root.find("WFProcesses") or root.find("Workflows")
        if wf_parent is not None:
            wfs = wf_parent.findall("WFProcess") or wf_parent.findall("Workflow")
            if wfs:
                build_workflows(doc, root, _next_sn(), maps,
                                render_dir=render_dir, font=pil_font)

    # eForms
    if "eforms" in sec:
        ef_parent = root.find("EForms")
        if ef_parent is not None and list(ef_parent):
            build_eforms(doc, root, _next_sn(), maps=maps,
                         render_dir=render_dir, font=pil_font)

    # Category Relationships
    if "relationships" in sec:
        build_relationship_diagram(doc, root, maps, categories, _next_sn(),
                                   render_dir=render_dir)

    # Folder / Object Hierarchy
    if "folder_structure" in sec and maps["folder_paths"]:
        build_folder_hierarchy(doc, root, maps, categories, _next_sn())

    # Keyword Dictionaries
    if "keyword_dicts" in sec and maps["kw_dicts"]:
        build_keyword_dicts(doc, maps, categories, _next_sn())

    # Queries / Reports / Stamps
    for key, items, title, id_field in [
        ("queries", queries, "Queries", "QueryTemplateNo"),
        ("stamps",  stamps,  "Stamps",  "StampNo"),
    ]:
        if key in sec and items:
            build_list_section(doc, items, _next_sn(), title, id_field)

    if "reports" in sec and reports:
        build_reports_section(doc, reports, _next_sn())

    # Script Inventory
    if "script_inventory" in sec:
        build_script_inventory(doc, profiles, root, _next_sn(), maps=maps,
                               ai_url=getattr(args, "ai_url", None),
                               ai_model=getattr(args, "ai_model", None),
                               ai_key=getattr(args, "ai_key", "lm-studio"))

    doc.save(args.output)
    size_kb = Path(args.output).stat().st_size // 1024
    print(f"Document generated: {Path(args.output).name}  ({size_kb:,} KB)")

    if _tmp_dir:
        shutil.rmtree(_tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xml",            help="Path to TheConfiguration-*.xml")
    ap.add_argument("-o", "--output", default="Therefore_Documentation.docx")
    ap.add_argument("--render-dir",   help="Cache rendered PNGs here")
    ap.add_argument("--data-dir",     help="REST-pulled JSON dir")
    ap.add_argument("--no-images",    action="store_true")
    ap.add_argument("--skip-eforms",  action="store_true")
    ap.add_argument("--body-only",    action="store_true",
                    help="Skip title page — output is content only for merging into a wrapper doc")
    ap.add_argument("--start-section", type=int, default=1, metavar="N",
                    help="First section number to use (default: 1, use higher when merging after existing sections)")
    ap.add_argument("--theme", help="Path to a YAML theme file")
    ap.add_argument("--format", choices=["png", "svg"], default="png",
                    help="Category image format (default: png)")
    ap.add_argument("--sections", default="",
                    help="Comma-separated section keys to include (default: all)")
    args = ap.parse_args()

    theme = load_theme(args.theme) if args.theme else None
    sections_list = [s.strip() for s in args.sections.split(",") if s.strip()] or None

    warnings = generate(
        args.xml, args.output,
        skip_eforms   = args.skip_eforms,
        no_images     = args.no_images,
        body_only     = args.body_only,
        render_dir    = getattr(args, "render_dir",    None),
        data_dir      = getattr(args, "data_dir",      None),
        start_section = getattr(args, "start_section", 1),
        theme         = theme,
        img_format    = args.format,
        sections      = sections_list,
    )
    for w in warnings:
        print(f"⚠  {w}")


if __name__ == "__main__":
    main()
