"""
CLI entry point for the FPL slide engine.

Usage:
    python generate.py "I need a vendor comparison slide for Salesforce vs CGI vs Maximo"
    python generate.py "Executive summary for the PD work management transformation" --archetype exec_summary
    python generate.py --list-archetypes
"""

import argparse
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Load .env if present
load_dotenv(Path(__file__).parent / ".env")

from slide_engine import SlideEngine
from slide_engine.compendium import list_archetypes, get_archetype_info


def main():
    parser = argparse.ArgumentParser(
        description="Generate BCG-style PowerPoint slides from natural language."
    )
    parser.add_argument("prompt", nargs="?", help="What the slide should be about")
    parser.add_argument("--archetype", "-a", help="Force a specific archetype (skip auto-selection)")
    parser.add_argument("--output", "-o", help="Output filename (without .pptx)")
    parser.add_argument("--list-archetypes", action="store_true", help="List all known archetypes")
    parser.add_argument("--output-dir", default="outputs", help="Output directory (default: outputs/)")
    args = parser.parse_args()

    if args.list_archetypes:
        print("\nAvailable archetypes:\n")
        for key in list_archetypes():
            info = get_archetype_info(key)
            print(f"  {key:<30} {info['description']}")
        print()
        return

    if not args.prompt:
        parser.print_help()
        sys.exit(1)

    engine = SlideEngine(output_dir=args.output_dir)
    result = engine.generate(
        prompt=args.prompt,
        archetype=args.archetype,
        output_name=args.output,
    )

    print("\n" + "─" * 60)
    print(result.summary())
    print("─" * 60)


if __name__ == "__main__":
    main()
