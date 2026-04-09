"""
LLM Agents (Phases 1, 3, 4, 7)
--------------------------------
Phase 1  - Intake:     select archetype from user prompt
Phase 3  - Architect:  generate structured slide content
Phase 4  - Critic:     review and score; loop until quality bar met
Phase 7  - Reviewer:   final narrative check; rewrite if needed

Uses claude-opus-4-6 with adaptive thinking throughout.
"""

import json
import os
import anthropic
from .compendium import describe_archetypes_for_llm

MODEL   = "claude-opus-4-6"
CLIENT  = None   # lazy-init so we don't fail on import if key is missing

BCG_DOCTRINE = """
## BCG Slide-Writing Doctrine

TITLES
- Action titles: state the insight, not the topic. "Salesforce leads on dispatch but lacks analytics" not "Vendor comparison"
- Max ~10 words. Sentence case. No period at end.
- If a subtitle exists, use it for the "so what" or context.

CONTENT
- Pyramid structure: lead with the conclusion, support below
- Each bullet = one idea. ~8–12 words per bullet. No full sentences.
- 3–5 bullets per column. Never more than 6.
- Use parallel construction across bullets in the same shape
- Numbers > words where possible ("$41M" not "forty-one million dollars")
- No orphan bullets (a single item in a group)

NARRATIVE
- Every shape should have a clear role in the slide story
- Title + body should be self-contained (readable without context)
- Avoid hedge language: "may", "could potentially", "it seems"

FORMATTING
- No bold, italics, or underline in body text (titles can be bold per template)
- Sentence case everywhere (not Title Case)
- Oxford comma
"""


def _get_client() -> anthropic.Anthropic:
    global CLIENT
    if CLIENT is None:
        key = os.environ.get("ANTHROPIC_API_KEY")
        if not key:
            raise EnvironmentError("ANTHROPIC_API_KEY not set. Create a .env file or export it.")
        CLIENT = anthropic.Anthropic(api_key=key)
    return CLIENT


# ---------------------------------------------------------------------------
# Phase 1 — Intake: select archetype
# ---------------------------------------------------------------------------

def select_archetype(user_prompt: str) -> str:
    """
    Given a free-text user prompt, return the best-matching archetype key.
    """
    archetypes_desc = describe_archetypes_for_llm()
    client = _get_client()

    response = client.messages.create(
        model=MODEL,
        max_tokens=256,
        thinking={"type": "adaptive"},
        messages=[{
            "role": "user",
            "content": f"""You are selecting a PowerPoint slide archetype.

Available archetypes:
{archetypes_desc}

User request: "{user_prompt}"

Reply with ONLY the archetype key (e.g. "exec_summary", "vendor_comparison").
No explanation."""
        }]
    )
    return _extract_text(response).strip().strip('"').strip("'")


# ---------------------------------------------------------------------------
# Phase 3 — Architect: generate content
# ---------------------------------------------------------------------------

def architect(user_prompt: str, archetype: str, spatial_summary: str) -> dict:
    """
    Generate structured slide content for the given archetype.
    Returns a dict mapping shape names/ph_types to text content.
    """
    client = _get_client()

    system = f"""You are a BCG slide-writing expert. Generate content for a single PowerPoint slide.
Output ONLY valid JSON — a flat object mapping shape roles to text content.
Shape role keys: "title", "body", or the shape name from the layout.
For multi-column layouts, use "body_1", "body_2", etc.

{BCG_DOCTRINE}"""

    prompt = f"""Slide type: {archetype}
User request: {user_prompt}

{spatial_summary}

Generate the slide content. Return JSON only, like:
{{
  "title": "Action-oriented title here",
  "body": "Line 1\\nLine 2\\nLine 3"
}}"""

    response = client.messages.create(
        model=MODEL,
        max_tokens=2048,
        thinking={"type": "adaptive"},
        system=system,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = _extract_text(response).strip()
    return _parse_json_response(raw)


# ---------------------------------------------------------------------------
# Phase 4 — Critic: score and refine
# ---------------------------------------------------------------------------

def critic(content: dict, archetype: str, spatial_summary: str, max_loops: int = 2) -> dict:
    """
    Score the content against BCG doctrine. If score < 8/10, rewrite and re-score.
    Returns the final (approved) content dict.
    """
    client = _get_client()

    system = f"""You are a rigorous BCG principal reviewing a slide for quality.
Score 1–10 against these criteria:
1. Action title (states the insight)
2. Bullet brevity (≤12 words each)
3. Parallel construction
4. Pyramid structure (conclusion first)
5. No hedge language

If score < 8, return improved content.
Output JSON only:
{{"score": 7, "issues": ["..."], "content": {{...improved content...}}}}
If score >= 8:
{{"score": 9, "issues": [], "content": null}}

{BCG_DOCTRINE}"""

    for attempt in range(max_loops):
        prompt = f"""Archetype: {archetype}
{spatial_summary}

Content to review:
{json.dumps(content, indent=2)}"""

        response = client.messages.create(
            model=MODEL,
            max_tokens=2048,
            thinking={"type": "adaptive"},
            system=system,
            messages=[{"role": "user", "content": prompt}]
        )

        result = _parse_json_response(_extract_text(response).strip())
        score = result.get("score", 10)

        if score >= 8 or result.get("content") is None:
            return content   # approved

        content = result["content"]   # use improved version

    return content   # return best after max_loops


# ---------------------------------------------------------------------------
# Phase 7 — Narrative Reviewer: final check
# ---------------------------------------------------------------------------

def narrative_review(content: dict, archetype: str) -> dict:
    """
    Final expert read of the slide. May rewrite if narrative is broken.
    Returns the final content dict.
    """
    client = _get_client()

    system = f"""You are a BCG MDP doing a final review of a slide before client delivery.
Check narrative coherence: does the title match the body? Is the story clear?
If the slide is ready, return: {{"approved": true, "content": null}}
If it needs a rewrite, return: {{"approved": false, "content": {{...rewritten...}}}}
Output JSON only.

{BCG_DOCTRINE}"""

    response = client.messages.create(
        model=MODEL,
        max_tokens=2048,
        thinking={"type": "adaptive"},
        system=system,
        messages=[{
            "role": "user",
            "content": f"Archetype: {archetype}\n\n{json.dumps(content, indent=2)}"
        }]
    )

    result = _parse_json_response(_extract_text(response).strip())
    if result.get("approved") is False and result.get("content"):
        return result["content"]
    return content


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _extract_text(response) -> str:
    for block in response.content:
        if block.type == "text":
            return block.text
    return ""


def _parse_json_response(raw: str) -> dict:
    """Parse JSON from LLM response, stripping markdown fences if present."""
    # Strip ```json ... ``` fences
    if raw.startswith("```"):
        lines = raw.split("\n")
        raw = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Try to extract first {...} block
        start = raw.find("{")
        end   = raw.rfind("}") + 1
        if start >= 0 and end > start:
            return json.loads(raw[start:end])
        return {}
