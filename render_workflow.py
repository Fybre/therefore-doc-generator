#!/usr/bin/env python3
"""
Therefore Workflow Renderer

Generates Mermaid flowchart syntax from Therefore workflow XML and renders
it to PNG via the Mermaid CLI (mmdc). Produces the same diagram style as
theconfiguration-processor HTML generator.

Requires: mmdc  (npm install -g @mermaid-js/mermaid-cli)

Usage:
    python render_workflow.py <input_xml> <output_dir>
"""

import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

try:
    from PIL import Image
    import numpy as np
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

_FIELD_REF_RE = re.compile(r'\[(-?\d+)\]')

# ---------------------------------------------------------------------------
# XML helpers (kept self-contained so this module works standalone)
# ---------------------------------------------------------------------------
def _get(elem, path):
    e = elem.find(path)
    return e.text if e is not None else None


def _localised(elem, tag):
    el = elem.find(tag)
    if el is None:
        return None
    s = el.find(".//TStr/T/S")
    if s is not None:
        return s.text
    s = el.find(".//S")
    return s.text if s is not None else None


def _name(elem):
    n = _localised(elem, "Name")
    return n if n else (_get(elem, "Name") or "")


def safe_fn(name):
    return re.sub(r"[^a-zA-Z0-9_-]", "_", name)


def safe_node_id(task_no: str) -> str:
    """Mermaid node IDs must be alphanumeric — strip the minus sign."""
    return "task" + str(task_no).replace("-", "n")


# ---------------------------------------------------------------------------
# Mermaid syntax generator  (matches theconfiguration-processor exactly)
# ---------------------------------------------------------------------------
def build_mermaid(wf_elem, field_no_map=None) -> str | None:
    """
    Generate Mermaid flowchart TD syntax for a workflow element.
    Returns None if the workflow has no tasks.
    """
    tasks_el = wf_elem.find("Tasks")
    if tasks_el is None:
        return None
    task_elems = tasks_el.findall("T") or tasks_el.findall("Task")
    if not task_elems:
        return None

    # Parse tasks
    tasks = {}
    for t in task_elems:
        tno = _get(t, "TaskNo")
        if not tno:
            continue
        tasks[tno] = {
            "name":        _name(t) or f"Task {tno}",
            "type_no":     int(_get(t, "Type") or 0),
            "transitions": [],
        }

    # Parse transitions
    for t in task_elems:
        tno = _get(t, "TaskNo")
        if not tno or tno not in tasks:
            continue
        trans_el = t.find("Transitions")
        if trans_el is None:
            continue
        for tr in trans_el.findall("TR"):
            to_no  = _get(tr, "TaskToNo") or ""
            action = _localised(tr, "ActionText") or _name(tr) or ""
            cond   = (_get(tr, "Condition") or "").strip()
            tasks[tno]["transitions"].append({
                "to":        to_no,
                "action":    action,
                "condition": cond,
            })

    lines = [
        "%%{init: {'flowchart': {'nodeSpacing': 30, 'rankSpacing': 50}}}%%",
        "flowchart TD",
    ]

    # Node declarations
    for tno, task in tasks.items():
        nid   = safe_node_id(tno)
        label = task["name"].replace('"', "'")
        ttype = task["type_no"]

        if ttype == 1:      # Start — stadium
            lines.append(f'    {nid}(["{label}"])')
        elif ttype == 2:    # End — double circle
            lines.append(f'    {nid}(("{label}"))')
        elif ttype == 4:    # Automatic — rectangle with (Auto) suffix
            lines.append(f'    {nid}["{label}\\n(Auto)"]')
        else:               # Manual / unknown — rectangle
            lines.append(f'    {nid}["{label}"]')

    # Transitions
    for tno, task in tasks.items():
        from_id = safe_node_id(tno)
        for tr in task["transitions"]:
            to_no  = tr["to"]
            if to_no not in tasks:
                continue
            to_id  = safe_node_id(to_no)
            action = tr["action"]
            cond   = tr["condition"]

            # Build edge label
            if cond:
                if field_no_map:
                    cond = _FIELD_REF_RE.sub(
                        lambda m: f"[{field_no_map.get(m.group(1), m.group(1))}]", cond
                    )
                short = cond[:30].replace('"', "'")
                if len(cond) > 30:
                    short += "..."
                label = f"{action}|IF: {short}|" if action else f"|IF: {short}|"
            else:
                label = action

            if label:
                label = label.replace('"', "'")
                lines.append(f'    {from_id} -->|"{label}"| {to_id}')
            else:
                lines.append(f'    {from_id} --> {to_id}')

    # Class definitions (same colours as theconfiguration-processor)
    lines.append("    classDef startNode fill:#d4edda,stroke:#28a745,stroke-width:2px")
    lines.append("    classDef endNode   fill:#f8d7da,stroke:#dc3545,stroke-width:2px")
    lines.append("    classDef manualNode fill:#cce5ff,stroke:#004085,stroke-width:2px")
    lines.append("    classDef autoNode  fill:#fff3cd,stroke:#856404,stroke-width:2px")

    # Class assignments
    for tno, task in tasks.items():
        nid   = safe_node_id(tno)
        ttype = task["type_no"]
        cls   = {1: "startNode", 2: "endNode", 3: "manualNode", 4: "autoNode"}.get(ttype)
        if cls:
            lines.append(f"    class {nid} {cls}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Rendering via mmdc
# ---------------------------------------------------------------------------
_MMDC = None  # resolved once

def _find_mmdc() -> str | None:
    global _MMDC
    if _MMDC is not None:
        return _MMDC
    for candidate in ["mmdc", "npx mmdc"]:
        try:
            result = subprocess.run(
                candidate.split() + ["--version"],
                capture_output=True, text=True, timeout=10,
            )
            if result.returncode == 0:
                _MMDC = candidate
                return _MMDC
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
    return None


def _autocrop(path: str, padding: int = 24) -> None:
    """Crop white margins from a rendered PNG in-place."""
    if not _PIL_AVAILABLE:
        return
    try:
        img = Image.open(path).convert("RGB")
        arr = np.array(img)
        mask = ~(arr == 255).all(axis=2)
        rows = np.any(mask, axis=1)
        cols = np.any(mask, axis=0)
        if not rows.any():
            return
        rmin, rmax = int(np.where(rows)[0][0]), int(np.where(rows)[0][-1])
        cmin, cmax = int(np.where(cols)[0][0]), int(np.where(cols)[0][-1])
        rmin = max(0, rmin - padding)
        rmax = min(arr.shape[0], rmax + padding)
        cmin = max(0, cmin - padding)
        cmax = min(arr.shape[1], cmax + padding)
        img.crop((cmin, rmin, cmax, rmax)).save(path)
    except Exception:
        pass  # non-fatal — leave original


def render_workflow(wf_elem, output_path: str, font=None, field_no_map=None) -> str | None:
    """
    Render a workflow element to a PNG via mmdc.
    `font` is accepted for API compatibility with the category renderer but unused.
    Returns output_path on success, None on failure.
    """
    mmdc = _find_mmdc()
    if not mmdc:
        return None

    mermaid_src = build_mermaid(wf_elem, field_no_map=field_no_map)
    if not mermaid_src:
        return None

    # Write Mermaid source to a temp file
    with tempfile.NamedTemporaryFile(mode="w", suffix=".mmd",
                                     delete=False, encoding="utf-8") as f:
        f.write(mermaid_src)
        tmp_mmd = f.name

    try:
        cmd = mmdc.split() + [
            "-i", tmp_mmd,
            "-o", output_path,
            "--width",           "900",
            "--height",          "6000",
            "--scale",           "2",
            "--backgroundColor", "white",
        ]
        _puppeteer_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), "puppeteer.json")
        if os.path.exists(_puppeteer_cfg):
            try:
                import json as _json
                _cfg = _json.loads(open(_puppeteer_cfg).read())
                _exe = _cfg.get("executablePath", "")
                # Only pass config if there's no executablePath, or it actually exists
                if not _exe or os.path.exists(_exe):
                    cmd += ["--puppeteerConfigFile", _puppeteer_cfg]
            except Exception:
                pass
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            print(f"  mmdc error: {result.stderr.strip()[:200]}", file=sys.stderr)
            return None
        if not os.path.exists(output_path):
            return None
        _autocrop(output_path)
        return output_path
    except subprocess.TimeoutExpired:
        print("  mmdc timeout", file=sys.stderr)
        return None
    finally:
        try:
            os.unlink(tmp_mmd)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Batch entry point
# ---------------------------------------------------------------------------
def render_all_workflows(input_xml: str, output_dir: str) -> list:
    if not _find_mmdc():
        print("Error: mmdc not found. Install with: npm install -g @mermaid-js/mermaid-cli")
        return []

    os.makedirs(output_dir, exist_ok=True)
    tree = ET.parse(input_xml)
    root = tree.getroot()

    wfp = root.find("WFProcesses")
    if wfp is None:
        wfp = root.find("Workflows")
    if wfp is None:
        print("No workflows found.")
        return []

    wfs = wfp.findall("WFProcess") or wfp.findall("Workflow")
    generated = []
    for wf in wfs:
        pno  = _get(wf, "ProcessNo") or _get(wf, "WFNo") or "x"
        name = _name(wf) or f"Workflow_{pno}"
        path = os.path.join(output_dir, f"{safe_fn(name)}_{pno}.png")
        result = render_workflow(wf, path)
        if result:
            generated.append(result)
            print(f"{name}: {path}")
        else:
            print(f"{name}: skipped")

    print(f"\nDone. {len(generated)} workflow image(s) in {output_dir}")
    return generated


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    render_all_workflows(sys.argv[1], sys.argv[2])
