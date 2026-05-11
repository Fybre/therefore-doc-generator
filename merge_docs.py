#!/usr/bin/env python3
"""
merge_docs.py — Insert a Therefore documentation body into a wrapper Word document.

In your wrapper document, add a line containing exactly:

    {{THEREFORE_CONTENT}}

wherever you want the content inserted (its own paragraph, nothing else on the line).
The script replaces that line with the full body content.

Usage:
    python merge_docs.py wrapper.docx body.docx output.docx
    python merge_docs.py wrapper.docx body.docx output.docx --placeholder "{{MY_SECTION}}"

Example workflow:
    # 1. Generate body-only content
    python build_doc.py TheConfiguration.xml -o body.docx --body-only

    # 2. Merge into your wrapper (which contains {{THEREFORE_CONTENT}} as a placeholder)
    python merge_docs.py "My Project Documentation.docx" body.docx output.docx
"""

import argparse
import copy
import sys


DEFAULT_PLACEHOLDER = "{{THEREFORE_CONTENT}}"

# XML namespaces
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _strip_heading_numbering(elements):
    """
    Remove <w:numPr> from heading-style paragraphs in the elements to insert.
    Without this, the wrapper document's auto-numbered heading list restarts
    at 1 for all inserted headings instead of continuing the existing sequence.
    """
    wP    = f"{{{_W_NS}}}p"
    wPPr  = f"{{{_W_NS}}}pPr"
    wPSty = f"{{{_W_NS}}}pStyle"
    wNPr  = f"{{{_W_NS}}}numPr"
    wVal  = f"{{{_W_NS}}}val"

    def _walk(el):
        if el.tag == wP:
            pPr = el.find(wPPr)
            if pPr is not None:
                ps = pPr.find(wPSty)
                if ps is not None and ps.get(wVal, "").startswith("Heading"):
                    numPr = pPr.find(wNPr)
                    if numPr is not None:
                        pPr.remove(numPr)
        for child in el:
            _walk(child)

    for el in elements:
        _walk(el)


def _copy_parts(body_doc, master):
    """
    Copy all non-external parts (images, charts, etc.) from body_doc into master.
    Returns a dict mapping old relationship IDs to new ones so XML references
    can be updated after the elements are inserted.
    """
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    rId_map = {}

    # Collect partnames already in master to avoid collisions
    existing = {str(p.partname) for p in master.part.package.iter_parts()}

    for rId, rel in list(body_doc.part.rels.items()):
        if rel.is_external:
            continue
        target = rel.target_part
        if not hasattr(target, "blob"):
            continue
        # Only copy image/media parts — skip numbering, styles, customXml etc.
        # which would conflict with the master document's own definitions.
        rel_tail = rel.reltype.split("/")[-1].lower()
        if rel_tail not in ("image", "media"):
            continue

        # Find a unique partname in master's package
        orig = str(target.partname)
        new_pn = orig
        if new_pn in existing:
            base, _, ext = orig.rpartition(".")
            counter = 1
            while new_pn in existing:
                new_pn = f"{base}_{counter}.{ext}"
                counter += 1

        existing.add(new_pn)
        new_part = Part(PackURI(new_pn), target.content_type,
                        target.blob, master.part.package)
        new_rId = master.part.relate_to(new_part, rel.reltype)
        rId_map[rId] = new_rId

    return rId_map


def _update_rids(element, rId_map):
    """
    Walk an XML element tree and replace any r:embed / r:id / r:link
    attribute values that appear in rId_map with their new counterparts.
    """
    if not rId_map:
        return
    attrs = [f"{{{_R_NS}}}{a}" for a in ("embed", "id", "link", "href")]
    for el in element.iter():
        for attr in attrs:
            if attr in el.attrib and el.attrib[attr] in rId_map:
                el.attrib[attr] = rId_map[el.attrib[attr]]


def merge(wrapper_path: str, body_path: str, output_path: str,
          placeholder: str = DEFAULT_PLACEHOLDER):
    from docx import Document

    master   = Document(wrapper_path)
    body_doc = Document(body_path)

    # Find the placeholder paragraph
    master_body    = master.element.body
    placeholder_el = None
    for para in master.paragraphs:
        if para.text.strip() == placeholder.strip():
            placeholder_el = para._element
            break

    if placeholder_el is None:
        print(f"Error: placeholder '{placeholder}' not found in {wrapper_path}")
        print("       Add a paragraph containing exactly that text where you want the content inserted.")
        sys.exit(1)

    # Copy all image/media parts and get the rId translation map
    rId_map = _copy_parts(body_doc, master)

    # Collect body elements to insert (skip final sectPr)
    elements_to_insert = [
        el for el in body_doc.element.body
        if not el.tag.endswith("}sectPr")
    ]

    # Strip auto-list-numbering from headings — prevents the wrapper's heading
    # list counter from restarting at 1 for all inserted headings.
    _strip_heading_numbering(elements_to_insert)

    # Find insertion index
    insert_idx = list(master_body).index(placeholder_el)

    # Insert in reverse order so index arithmetic stays correct
    for el in reversed(elements_to_insert):
        copied = copy.deepcopy(el)
        _update_rids(copied, rId_map)
        master_body.insert(insert_idx, copied)

    # Remove the placeholder paragraph
    master_body.remove(placeholder_el)

    master.save(output_path)
    print(f"Inserted {len(elements_to_insert)} elements ({len(rId_map)} media parts copied).")
    print(f"Saved: {output_path}")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("wrapper",       help="Wrapper .docx containing the placeholder")
    ap.add_argument("body",          help="Body .docx generated with --body-only")
    ap.add_argument("output",        help="Output .docx path")
    ap.add_argument("--placeholder", default=DEFAULT_PLACEHOLDER,
                    help=f"Placeholder text to replace (default: {DEFAULT_PLACEHOLDER})")
    args = ap.parse_args()

    merge(args.wrapper, args.body, args.output, placeholder=args.placeholder)


if __name__ == "__main__":
    main()
