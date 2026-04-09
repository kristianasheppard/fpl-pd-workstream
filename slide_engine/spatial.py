"""
Spatial Awareness Engine (Phase 2)
-----------------------------------
Analyzes the shape layout of a template slide to give the Architect LLM
a precise understanding of available content zones, text capacities, and
spatial relationships between shapes.

EMU (English Metric Units): 914400 EMU = 1 inch, 12700 EMU = 1 pt
"""

from lxml import etree
from .cloner import NS, _tag, _describe_shape

# Slide dimensions (standard widescreen 16:9)
SLIDE_WIDTH_EMU  = 12192000
SLIDE_HEIGHT_EMU = 6858000

# Rough character capacity estimate: EMU to approximate chars
# Assumes ~12pt body text, Arial, standard line height
EMU_PER_CHAR_WIDTH  = 120000   # ~9pt char width at 12pt
EMU_PER_LINE_HEIGHT = 200000   # ~16pt line height at 12pt


def analyze_slide(slide_xml: str | bytes) -> dict:
    """
    Parse a slide's OOXML and return a spatial manifest for the LLM.
    Returns a dict with:
      - shapes: list of shape descriptors with capacity estimates
      - layout_summary: human-readable description
    """
    if isinstance(slide_xml, str):
        slide_xml = slide_xml.encode("utf-8")

    tree = etree.fromstring(slide_xml)
    shapes = []

    for sp in tree.iter(_tag("p", "sp")):
        desc = _describe_shape(sp)
        if desc:
            desc["capacity"] = _estimate_capacity(desc)
            desc["position_label"] = _position_label(desc)
            shapes.append(desc)

    # Sort spatially: top-to-bottom, left-to-right
    shapes.sort(key=lambda s: (s["y"], s["x"]))

    return {
        "shapes": shapes,
        "layout_summary": _summarize_layout(shapes),
    }


def format_for_prompt(spatial: dict) -> str:
    """
    Format spatial analysis as a compact string for inclusion in LLM prompts.
    """
    lines = ["## Slide Layout\n"]
    for s in spatial["shapes"]:
        role = s["ph_type"] or "content"
        pos  = s["position_label"]
        cap  = s["capacity"]
        curr = s["text"][:60] + "…" if len(s["text"]) > 60 else s["text"]
        lines.append(
            f"- Shape: \"{s['name']}\" | Role: {role} | Position: {pos} | "
            f"Capacity: ~{cap['chars']} chars / {cap['lines']} lines"
            + (f" | Current text: \"{curr}\"" if curr else "")
        )
    lines.append(f"\n{spatial['layout_summary']}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _estimate_capacity(shape: dict) -> dict:
    """Estimate rough character and line capacity for a shape."""
    cx, cy = shape["cx"], shape["cy"]
    if cx <= 0 or cy <= 0:
        return {"chars": 200, "lines": 5}

    # Padding: assume ~5% margin inside shape
    usable_cx = int(cx * 0.90)
    usable_cy = int(cy * 0.90)

    chars_per_line = max(1, usable_cx // EMU_PER_CHAR_WIDTH)
    num_lines      = max(1, usable_cy // EMU_PER_LINE_HEIGHT)
    total_chars    = chars_per_line * num_lines

    return {
        "chars": total_chars,
        "lines": num_lines,
        "chars_per_line": chars_per_line,
    }


def _position_label(shape: dict) -> str:
    """Return a human-readable position label (top/bottom, left/center/right)."""
    x, y, cx, cy = shape["x"], shape["y"], shape["cx"], shape["cy"]
    center_x = x + cx // 2
    center_y = y + cy // 2

    v = "top" if center_y < SLIDE_HEIGHT_EMU * 0.35 else \
        "bottom" if center_y > SLIDE_HEIGHT_EMU * 0.65 else "middle"

    h = "left" if center_x < SLIDE_WIDTH_EMU * 0.33 else \
        "right" if center_x > SLIDE_WIDTH_EMU * 0.67 else "center"

    return f"{v}-{h}"


def _summarize_layout(shapes: list[dict]) -> str:
    """Generate a one-line layout description."""
    ph_types = [s["ph_type"] for s in shapes if s["ph_type"]]
    content_count = sum(1 for s in shapes if s["ph_type"] in ("body", None) and s["cx"] > 1_000_000)

    parts = []
    if "title" in ph_types or "ctrTitle" in ph_types:
        parts.append("title")
    if content_count == 1:
        parts.append("1 content zone")
    elif content_count == 2:
        parts.append("2-column layout")
    elif content_count >= 3:
        parts.append(f"{content_count}-column layout")
    if "subTitle" in ph_types:
        parts.append("subtitle")

    return "Layout: " + " + ".join(parts) if parts else "Layout: custom"
