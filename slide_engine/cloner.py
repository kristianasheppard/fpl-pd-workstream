"""
OOXML Slide Cloner
------------------
Core engine: takes a source .pptx and a slide number, clones that slide into
a new single-slide .pptx, and injects replacement text at the shape level.

Key principle (from AJ Tella's breakthrough): unzip -> deconstruct OOXML ->
reconstruct. If you can clone a slide, everything else flows from there.
"""

import copy
import zipfile
from pathlib import Path
from lxml import etree

# PowerPoint XML namespace map
NS = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p":   "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct":  "http://schemas.openxmlformats.org/package/2006/content-types",
}

def _tag(ns_key: str, local: str) -> str:
    return f"{{{NS[ns_key]}}}{local}"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_shape_map(pptx_path: str | Path, slide_number: int) -> list[dict]:
    """
    Return a list of shape descriptors for the given slide (1-indexed).
    Each descriptor: {id, name, ph_type, text, x, y, cx, cy}
    Used by the spatial awareness phase.
    """
    pptx_path = Path(pptx_path)
    with zipfile.ZipFile(pptx_path) as z:
        slide_names = _get_slide_names(z)
        if slide_number < 1 or slide_number > len(slide_names):
            raise ValueError(f"Slide {slide_number} out of range (1–{len(slide_names)})")
        xml_bytes = z.read(slide_names[slide_number - 1])

    tree = etree.fromstring(xml_bytes)
    shapes = []
    for sp in tree.iter(_tag("p", "sp")):
        shape = _describe_shape(sp)
        if shape:
            shapes.append(shape)
    return shapes


def clone_slide(
    source_pptx: str | Path,
    slide_number: int,
    text_map: dict[str, str],
    output_path: str | Path,
) -> Path:
    """
    Clone slide `slide_number` from `source_pptx` into a new .pptx at
    `output_path`, replacing text in shapes per `text_map`.

    text_map keys can be:
      - placeholder type: "title", "body", "subTitle", "dt", "ftr", "sldNum"
      - shape name (partial match, case-insensitive): "Content Placeholder 2"

    Returns the output_path as a Path.
    """
    source_pptx = Path(source_pptx)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source_pptx, "r") as zin:
        all_files = {name: zin.read(name) for name in zin.namelist()}

    slide_names = _get_slide_names_from_dict(all_files)
    if slide_number < 1 or slide_number > len(slide_names):
        raise ValueError(f"Slide {slide_number} out of range (1–{len(slide_names)})")

    target = slide_names[slide_number - 1]
    target_base = target.split("/")[-1]               # e.g. "slide3.xml"
    to_remove = set(slide_names) - {target}

    # Also remove rels for removed slides
    for s in list(to_remove):
        to_remove.add(f"ppt/slides/_rels/{s.split('/')[-1]}.rels")

    for path in to_remove:
        all_files.pop(path, None)

    # Trim presentation.xml sldIdLst to only the target slide
    all_files["ppt/presentation.xml"] = _trim_presentation(
        all_files["ppt/presentation.xml"],
        all_files.get("ppt/_rels/presentation.xml.rels", b""),
        target_base,
    )

    # Trim [Content_Types].xml
    all_files["[Content_Types].xml"] = _trim_content_types(
        all_files["[Content_Types].xml"], to_remove
    )

    # Inject text into the target slide
    all_files[target] = _inject_text(all_files[target], text_map)

    # Prune orphaned media (keep only media referenced by the target slide's rels)
    target_rels_path = f"ppt/slides/_rels/{target_base}.rels"
    referenced_media = _get_referenced_media(all_files.get(target_rels_path, b""), target)
    for path in list(all_files):
        if path.startswith("ppt/media/") and path not in referenced_media:
            del all_files[path]

    # Write new zip
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)

    return output_path


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_slide_names(z: zipfile.ZipFile) -> list[str]:
    return sorted(
        n for n in z.namelist()
        if n.startswith("ppt/slides/slide") and n.endswith(".xml") and "_rels" not in n
    )


def _get_slide_names_from_dict(files: dict) -> list[str]:
    return sorted(
        n for n in files
        if n.startswith("ppt/slides/slide") and n.endswith(".xml") and "_rels" not in n
    )


def _trim_presentation(prs_xml: bytes, rels_xml: bytes, target_base: str) -> bytes:
    """Remove all sldId entries except the target from presentation.xml."""
    tree = etree.fromstring(prs_xml)
    p_ns = NS["p"]
    r_ns = NS["r"]
    pkg_ns = NS["pkg"]

    # Find rId for target_base in presentation rels
    target_rid = None
    if rels_xml:
        rels_tree = etree.fromstring(rels_xml)
        for rel in rels_tree.iter(f"{{{pkg_ns}}}Relationship"):
            target_val = rel.get("Target", "")
            if target_val.endswith(target_base) or target_val == f"slides/{target_base}":
                target_rid = rel.get("Id")
                break

    sld_id_lst = tree.find(f"{{{p_ns}}}sldIdLst")
    if sld_id_lst is not None and target_rid:
        for sld_id in list(sld_id_lst):
            rid = sld_id.get(f"{{{r_ns}}}id")
            if rid != target_rid:
                sld_id_lst.remove(sld_id)

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _trim_content_types(ct_xml: bytes, removed_paths: set[str]) -> bytes:
    """Remove Override entries for removed slides from [Content_Types].xml."""
    tree = etree.fromstring(ct_xml)
    ct_ns = NS["ct"]
    for override in list(tree.findall(f"{{{ct_ns}}}Override")):
        part = override.get("PartName", "").lstrip("/")
        if part in removed_paths:
            tree.remove(override)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _get_referenced_media(rels_xml: bytes, slide_path: str) -> set[str]:
    """Return the set of ppt/media/... paths referenced by a slide's rels file."""
    if not rels_xml:
        return set()
    slide_dir = "/".join(slide_path.split("/")[:-1])  # "ppt/slides"
    tree = etree.fromstring(rels_xml)
    pkg_ns = NS["pkg"]
    media = set()
    for rel in tree.iter(f"{{{pkg_ns}}}Relationship"):
        target = rel.get("Target", "")
        if "../media/" in target:
            # Resolve relative path: ppt/slides/../media/imageX.png → ppt/media/imageX.png
            resolved = "/".join(slide_dir.split("/")[:-1]) + "/media/" + target.split("../media/")[-1]
            media.add(resolved)
    return media


def _describe_shape(sp) -> dict | None:
    """Extract descriptor from a <p:sp> element."""
    p_ns = NS["p"]
    a_ns = NS["a"]

    nv = sp.find(f"{{{p_ns}}}nvSpPr")
    if nv is None:
        return None

    cnv = nv.find(f"{{{p_ns}}}cNvPr")
    shape_id   = int(cnv.get("id", 0)) if cnv is not None else 0
    shape_name = cnv.get("name", "") if cnv is not None else ""

    nv_pr = nv.find(f"{{{p_ns}}}nvPr")
    ph_type = None
    if nv_pr is not None:
        ph = nv_pr.find(f"{{{p_ns}}}ph")
        if ph is not None:
            ph_type = ph.get("type", "body")

    # Position / size
    sp_pr = sp.find(f"{{{p_ns}}}spPr")
    x = y = cx = cy = 0
    if sp_pr is not None:
        xfrm = sp_pr.find(f"{{{a_ns}}}xfrm")
        if xfrm is not None:
            off = xfrm.find(f"{{{a_ns}}}off")
            ext = xfrm.find(f"{{{a_ns}}}ext")
            if off is not None:
                x, y = int(off.get("x", 0)), int(off.get("y", 0))
            if ext is not None:
                cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))

    # Text content
    tx_body = sp.find(f"{{{p_ns}}}txBody")
    text_lines = []
    if tx_body is not None:
        for para in tx_body.findall(f"{{{a_ns}}}p"):
            line = "".join(
                t.text or "" for t in para.iter(f"{{{a_ns}}}t")
            )
            if line:
                text_lines.append(line)

    return {
        "id":      shape_id,
        "name":    shape_name,
        "ph_type": ph_type,
        "text":    "\n".join(text_lines),
        "x": x, "y": y, "cx": cx, "cy": cy,
    }


def _inject_text(slide_xml: bytes, text_map: dict[str, str]) -> bytes:
    """Replace text in shapes matching keys in text_map."""
    tree = etree.fromstring(slide_xml)
    p_ns = NS["p"]
    a_ns = NS["a"]

    for sp in tree.iter(f"{{{p_ns}}}sp"):
        shape = _describe_shape(sp)
        if not shape:
            continue

        new_text = _find_match(text_map, shape["ph_type"], shape["name"])
        if new_text is None:
            continue

        tx_body = sp.find(f"{{{p_ns}}}txBody")
        if tx_body is not None:
            _set_text(tx_body, new_text, a_ns)

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _find_match(text_map: dict, ph_type: str | None, shape_name: str) -> str | None:
    """Find the best matching key in text_map for this shape."""
    for key, val in text_map.items():
        if key == ph_type:
            return val
        if key.lower() in shape_name.lower() or shape_name.lower() in key.lower():
            return val
    return None


def _set_text(tx_body, new_text: str, a_ns: str):
    """Replace all text in a txBody with new_text, preserving run formatting."""
    paras = tx_body.findall(f"{{{a_ns}}}p")
    if not paras:
        return

    # Capture the first paragraph as format template
    template_para = copy.deepcopy(paras[0])

    # Get template run properties (font, size, bold, etc.)
    tmpl_run = template_para.find(f"{{{a_ns}}}r")
    tmpl_rpr = tmpl_run.find(f"{{{a_ns}}}rPr") if tmpl_run is not None else None

    # Remove all existing paragraphs from tx_body
    for p in paras:
        tx_body.remove(p)

    # Insert new paragraphs for each line
    body_pr = tx_body.find(f"{{{a_ns}}}bodyPr")
    if body_pr is not None:
        insert_idx = list(tx_body).index(body_pr) + 1
    else:
        insert_idx = len(tx_body)

    lines = new_text.split("\n") if new_text else [""]
    for i, line in enumerate(lines):
        new_p = copy.deepcopy(template_para)
        # Remove existing runs/breaks
        for child in list(new_p):
            if child.tag in (f"{{{a_ns}}}r", f"{{{a_ns}}}br"):
                new_p.remove(child)

        if line.strip():
            r = etree.SubElement(new_p, f"{{{a_ns}}}r")
            if tmpl_rpr is not None:
                r.insert(0, copy.deepcopy(tmpl_rpr))
            t = etree.SubElement(r, f"{{{a_ns}}}t")
            t.text = line
            if line != line.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        tx_body.insert(insert_idx + i, new_p)
