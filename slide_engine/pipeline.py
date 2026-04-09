"""
SlideEngine — 8-Phase Pipeline Orchestrator
--------------------------------------------
Implements AJ Tella's multi-agent pipeline:

  1. Intake          → parse prompt, select archetype
  2. Spatial         → analyze template layout
  3. Architect LLM   → generate content within constraints
  4. Critic LLM      → review, loop until quality bar met
  5. Edit Mapping    → map content keys to shape IDs
  6. OOXML Apply     → inject content into clone
  7. Reviewer LLM    → final narrative check
  8. Structural QA   → verify output is a valid .pptx

Usage:
    engine = SlideEngine()
    result = engine.generate("I need a vendor comparison slide for Salesforce vs CGI")
    print(result.output_path)
"""

import time
from dataclasses import dataclass, field
from pathlib import Path

from .compendium import get_source_pptx, get_primary_slide_number
from .spatial    import analyze_slide, format_for_prompt
from .cloner     import clone_slide, get_shape_map
from .agents     import select_archetype, architect, critic, narrative_review
from .compendium import get_template_xml


OUTPUT_DIR = Path(__file__).parent.parent / "outputs"


@dataclass
class SlideResult:
    archetype:    str
    content:      dict
    output_path:  Path
    phase_log:    list[str] = field(default_factory=list)
    elapsed_sec:  float = 0.0

    def summary(self) -> str:
        lines = [
            f"Archetype:  {self.archetype}",
            f"Output:     {self.output_path}",
            f"Time:       {self.elapsed_sec:.1f}s",
            "Phases:     " + " → ".join(self.phase_log),
            "",
            "Content:",
        ]
        for key, val in self.content.items():
            preview = val[:80].replace("\n", " | ") + ("…" if len(val) > 80 else "")
            lines.append(f"  {key}: {preview}")
        return "\n".join(lines)


class SlideEngine:
    """
    Main entry point for slide generation.
    """

    def __init__(self, output_dir: str | Path | None = None):
        self.output_dir = Path(output_dir) if output_dir else OUTPUT_DIR
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def generate(
        self,
        prompt: str,
        archetype: str | None = None,
        output_name: str | None = None,
    ) -> SlideResult:
        """
        Generate a slide from a natural-language prompt.

        Args:
            prompt:      What the slide should say / be about
            archetype:   Force a specific archetype (skips Phase 1 if set)
            output_name: Output filename (without .pptx); auto-generated if None
        """
        t0 = time.time()
        log = []

        # ── Phase 1: Intake ──────────────────────────────────────────────
        print(f"[1/8] Intake: selecting archetype…")
        if archetype is None:
            archetype = select_archetype(prompt)
        log.append(f"1:intake({archetype})")
        print(f"      → {archetype}")

        # ── Phase 2: Spatial Awareness ───────────────────────────────────
        print(f"[2/8] Spatial: analyzing template layout…")
        template_xml = get_template_xml(archetype)
        spatial      = analyze_slide(template_xml)
        spatial_str  = format_for_prompt(spatial)
        log.append("2:spatial")
        print(f"      → {spatial['layout_summary']}")

        # ── Phase 3: Architect LLM ───────────────────────────────────────
        print(f"[3/8] Architect: generating content…")
        content = architect(prompt, archetype, spatial_str)
        log.append("3:architect")
        print(f"      → {len(content)} shapes filled")

        # ── Phase 4: Critic LLM ──────────────────────────────────────────
        print(f"[4/8] Critic: reviewing quality…")
        content = critic(content, archetype, spatial_str)
        log.append("4:critic")

        # ── Phase 5: Edit Mapping ────────────────────────────────────────
        print(f"[5/8] Edit mapping: resolving shape keys…")
        text_map = _resolve_text_map(content, spatial["shapes"])
        log.append("5:mapping")
        print(f"      → {len(text_map)} shapes mapped")

        # ── Phase 6: OOXML Application ───────────────────────────────────
        print(f"[6/8] OOXML: cloning and injecting content…")
        source_pptx  = get_source_pptx(archetype)
        slide_number = get_primary_slide_number(archetype)
        out_name     = output_name or f"{archetype}_{int(t0)}.pptx"
        output_path  = self.output_dir / out_name

        clone_slide(source_pptx, slide_number, text_map, output_path)
        log.append("6:ooxml")
        print(f"      → {output_path}")

        # ── Phase 7: Narrative Review ────────────────────────────────────
        print(f"[7/8] Reviewer: final narrative check…")
        content = narrative_review(content, archetype)
        # If reviewer changed content, re-apply to output
        revised_map = _resolve_text_map(content, spatial["shapes"])
        if revised_map != text_map:
            clone_slide(source_pptx, slide_number, revised_map, output_path)
            print(f"      → content revised by reviewer")
        log.append("7:reviewer")

        # ── Phase 8: Structural QA ───────────────────────────────────────
        print(f"[8/8] QA: validating output…")
        qa_ok, qa_msg = _structural_qa(output_path)
        log.append(f"8:qa({'ok' if qa_ok else 'warn'})")
        if not qa_ok:
            print(f"      ⚠ {qa_msg}")
        else:
            print(f"      → valid .pptx")

        elapsed = time.time() - t0
        result = SlideResult(
            archetype=archetype,
            content=content,
            output_path=output_path,
            phase_log=log,
            elapsed_sec=elapsed,
        )
        print(f"\nDone in {elapsed:.1f}s → {output_path}\n")
        return result


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _resolve_text_map(content: dict, shapes: list[dict]) -> dict[str, str]:
    """
    Map content keys (from LLM) back to shape names/ph_types the cloner accepts.
    Handles "body_1" / "body_2" → numbered body shapes by position.
    """
    text_map = {}
    body_shapes = [s for s in shapes if s["ph_type"] in ("body", None) and s["cx"] > 1_000_000]
    body_shapes.sort(key=lambda s: s["x"])   # left-to-right order

    for key, val in content.items():
        if not val:
            continue
        if key in ("title", "ctrTitle", "subTitle", "dt", "ftr", "sldNum"):
            text_map[key] = val
        elif key == "body":
            text_map["body"] = val
        elif key.startswith("body_"):
            # body_1 → first body shape, body_2 → second, etc.
            idx = int(key.split("_")[1]) - 1
            if 0 <= idx < len(body_shapes):
                text_map[body_shapes[idx]["name"]] = val
        else:
            # Pass through as shape name
            text_map[key] = val

    return text_map


def _structural_qa(output_path: Path) -> tuple[bool, str]:
    """Quick validation: is the .pptx a valid zip with the expected structure?"""
    import zipfile
    try:
        with zipfile.ZipFile(output_path) as z:
            names = z.namelist()
            has_prs = "ppt/presentation.xml" in names
            has_slide = any(n.startswith("ppt/slides/slide") and n.endswith(".xml") for n in names)
            if not has_prs:
                return False, "Missing ppt/presentation.xml"
            if not has_slide:
                return False, "No slides found in output"
        return True, "ok"
    except Exception as e:
        return False, str(e)
