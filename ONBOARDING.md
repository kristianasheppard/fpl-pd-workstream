# FPL PD Workstream — Onboarding Guide
**For: New team members joining the fpl-pd-workstream repo**

---

## What This Is

This repo contains a slide generation engine built for the FPL Power Delivery Work Management transformation project. It uses Claude (Anthropic's AI) to generate BCG-style PowerPoint slides from natural language prompts, pulling from a library of real FPL SteerCo slide templates.

You can say something like:
> *"Build me a vendor comparison slide for Salesforce vs CGI with a recommendation for Salesforce"*

...and get a properly formatted .pptx back in under 2 minutes.

---

## Step 1 — Get Repo Access

1. Create a free GitHub account at [github.com](https://github.com) if you don't have one
2. Send your GitHub username to **Kristian Sheppard** — he'll add you as a collaborator
3. Once added, clone the repo:
   ```bash
   git clone https://github.com/kristianasheppard/fpl-pd-workstream.git
   cd fpl-pd-workstream
   ```

---

## Step 2 — Python Setup

You need Python 3.11+. Check with:
```bash
python --version
```

If you don't have it, download from [python.org](https://python.org) or install [Miniconda](https://docs.conda.io/en/latest/miniconda.html).

Install dependencies:
```bash
pip install -r requirements.txt
```

This installs three packages: `anthropic` (the AI API), `lxml` (XML parser), and `python-dotenv` (env file loader).

---

## Step 3 — Get an API Key

The engine calls Claude via Anthropic's API. You need your own key.

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign up / log in
3. Navigate to **API Keys** and create a new key
4. Copy the key (it starts with `sk-ant-...`)

Create a `.env` file in the repo root:
```bash
cp .env.example .env
```

Open `.env` and paste your key:
```
ANTHROPIC_API_KEY=sk-ant-...your key here...
```

> **Important:** Never commit your `.env` file. It is already listed in `.gitignore`.

---

## Step 4 — Source Files (Required)

The engine clones slides from two source .pptx files that live in the BCG SharePoint. You need both downloaded locally:

| File | Path on SharePoint |
|------|--------------------|
| `260217_PD Work Management SteerCo Compendium.pptx` | FPL-10019021 > 000 - Phase 1 Reference > 05. Final deliverables > 16 - Final Deliverables |
| `FPL PD Work Management Final Deliverables_vShare.pptx` | Same folder |

Once downloaded, open `slide_engine/compendium.py` and update the `SOURCE_FILES` paths to match wherever you saved them:

```python
SOURCE_FILES = {
    "compendium": r"C:\path\to\your\260217_PD Work Management SteerCo Compendium.pptx",
    "final":      r"C:\path\to\your\FPL PD Work Management Final Deliverables_vShare.pptx",
}
```

---

## Step 5 — Run the Test

```bash
python test_run.py
```

This generates a 4-slide deck (cover, exec summary, findings, next steps) and saves it to:
```
outputs/test_deck/FPL_PD_WMS_Phase2_Kickoff.pptx
```

It takes about 4-5 minutes (the AI is doing real work on each slide). Open the file in PowerPoint when done.

---

## Step 6 — Generate Your Own Slides

### Single slide
```bash
python generate.py "Vendor comparison: Salesforce vs CGI vs Maximo. Salesforce wins on dispatch, mobility, and integration. CGI wins on cost."
```

### Force a specific archetype
```bash
python generate.py "Your prompt here" --archetype vendor_comparison
```

### See all available archetypes
```bash
python generate.py --list-archetypes
```

### From Python directly
```python
from slide_engine import SlideEngine

engine = SlideEngine(output_dir="outputs/my_deck")

# Single slide
result = engine.generate(
    prompt="Key decisions needed from leadership: ...",
    archetype="key_decisions"
)
print(result.output_path)

# Full deck
deck = engine.generate_deck(
    slides_spec=[
        {"archetype": "cover",        "prompt": "My deck title..."},
        {"archetype": "exec_summary", "prompt": "Key messages..."},
        {"archetype": "next_steps",   "prompt": "Action items..."},
    ],
    output_name="my_deck.pptx"
)
print(deck.output_path)
```

---

## Available Archetypes

| Archetype | Use for |
|-----------|---------|
| `cover` | Title slide |
| `exec_summary` | Opening summary with key messages |
| `agenda` | Meeting agenda |
| `key_decisions` | Decision log / asks from leadership |
| `for_discussion` | Discussion framing slide |
| `current_state_assessment` | As-is assessment |
| `analysis_findings` | Data findings, research output |
| `business_case` | NPV, ROI, financials |
| `vendor_comparison` | Side-by-side vendor evaluation |
| `north_star_vision` | Future state / vision |
| `process_workflow` | Process maps, swim lanes |
| `next_steps` | Actions, owners, dates |
| `section_break` | Section divider |

---

## Questions?

Reach out to **Kristian Sheppard** or open an issue in the repo.
