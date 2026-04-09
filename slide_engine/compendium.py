"""
Slide Compendium
----------------
Loads and manages the archetype catalog. Provides archetype selection
(either explicit or LLM-assisted) and template XML retrieval.

Adding new archetypes: add an entry to catalog.json and drop the slide XML
in templates/. New slides should be mapped to an existing bucket or a new
bucket is proposed.
"""

import json
from pathlib import Path

COMPENDIUM_DIR = Path(r"C:\Users\sheppard kristian\.claude\skills\slide-compendium")
CATALOG_PATH   = COMPENDIUM_DIR / "catalog.json"
TEMPLATES_DIR  = COMPENDIUM_DIR / "templates"

SOURCE_FILES = {
    "compendium": (
        r"C:\Users\sheppard kristian\The Boston Consulting Group, Inc"
        r"\FPL-10019021-FPL Power Delivery WMS - Documents"
        r"\000 - Phase 1 Reference\05. Final deliverables"
        r"\16 - Final Deliverables\260217_PD Work Management SteerCo Compendium.pptx"
    ),
    "final": (
        r"C:\Users\sheppard kristian\The Boston Consulting Group, Inc"
        r"\FPL-10019021-FPL Power Delivery WMS - Documents"
        r"\000 - Phase 1 Reference\05. Final deliverables"
        r"\16 - Final Deliverables\FPL PD Work Management Final Deliverables_vShare.pptx"
    ),
}


def load_catalog() -> dict:
    """Load the full archetype catalog from catalog.json."""
    with open(CATALOG_PATH) as f:
        return json.load(f)


def list_archetypes() -> list[str]:
    """Return the list of known archetype keys."""
    return list(load_catalog().keys())


def get_archetype_info(archetype: str) -> dict:
    """Return the catalog entry for a single archetype."""
    catalog = load_catalog()
    if archetype not in catalog:
        raise ValueError(f"Unknown archetype '{archetype}'. Known: {list(catalog)}")
    return catalog[archetype]


def get_template_xml(archetype: str) -> str:
    """Return the raw OOXML string for the primary template of an archetype."""
    xml_path = TEMPLATES_DIR / f"{archetype}.xml"
    if not xml_path.exists():
        raise FileNotFoundError(f"Template XML not found: {xml_path}")
    return xml_path.read_text(encoding="utf-8", errors="replace")


def get_source_pptx(archetype: str) -> Path:
    """Return the path to the source .pptx for the primary template."""
    info = get_archetype_info(archetype)
    file_key = info["primary"]["file"]
    return Path(SOURCE_FILES[file_key])


def get_primary_slide_number(archetype: str) -> int:
    """Return the 1-indexed slide number for the primary template."""
    return get_archetype_info(archetype)["primary"]["slide"]


def describe_archetypes_for_llm() -> str:
    """Return a compact description of all archetypes for use in LLM prompts."""
    catalog = load_catalog()
    lines = []
    for key, info in catalog.items():
        primary = info["primary"]
        lines.append(
            f"- {key}: {info['description']} "
            f"(primary: {primary['file']} slide {primary['slide']}, \"{primary['title']}\")"
        )
    return "\n".join(lines)
