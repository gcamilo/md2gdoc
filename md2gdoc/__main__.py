"""CLI entry point for md2gdoc."""

import argparse
import sys
from pathlib import Path

from . import __version__
from .engine import convert
from .styles import get_style, list_styles


def main():
    parser = argparse.ArgumentParser(
        prog="md2gdoc",
        description="Convert Markdown to formatted Google Docs. No pandoc.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  md2gdoc input.md                           Create new doc (no styling)
  md2gdoc input.md --style ib                IB pitch-book formatting
  md2gdoc input.md --style clean             Clean, minimal formatting
  md2gdoc input.md --style corporate         Corporate blue theme
  md2gdoc input.md --doc-id DOC_ID           Overwrite existing doc
  md2gdoc input.md --style ib --title "Q4"   Custom title + IB style
  md2gdoc input.md --style-file my.json      Custom style from JSON
  md2gdoc --list-styles                      Show available style presets
  md2gdoc --stdin --style ib                 Read markdown from stdin

credentials:
  Set GOOGLE_TOKEN_PATH env var, or place token at:
    ~/.config/md2gdoc/token.json
    ~/.config/google-docs-mcp/token.json
        """,
    )

    parser.add_argument("input", nargs="?", help="Markdown file path")
    parser.add_argument("--doc-id", help="Existing Google Doc ID to overwrite")
    parser.add_argument("--title", help="Document title (defaults to first H1)")
    parser.add_argument("--style", default=None,
                        help="Style preset: ib, clean, corporate, none (default: none)")
    parser.add_argument("--style-file", help="Path to custom style JSON file")
    parser.add_argument("--credentials", help="Path to Google OAuth token JSON")
    parser.add_argument("--no-diagrams", action="store_true",
                        help="Skip diagram embedding from Google Slides")
    parser.add_argument("--stdin", action="store_true",
                        help="Read markdown from stdin instead of file")
    parser.add_argument("--quiet", "-q", action="store_true",
                        help="Suppress progress output")
    parser.add_argument("--list-styles", action="store_true",
                        help="List available style presets and exit")
    parser.add_argument("--version", "-v", action="version",
                        version=f"md2gdoc {__version__}")

    args = parser.parse_args()

    if args.list_styles:
        print("Available style presets:\n")
        for name, desc in list_styles():
            print(f"  {name:12s}  {desc}")
        print(f"\n  {'none':12s}  No formatting (content only)")
        print("\nUse --style-file for custom styles (JSON).")
        sys.exit(0)

    # Read input
    if args.stdin:
        md_text = sys.stdin.read()
    elif args.input:
        md_path = Path(args.input)
        if not md_path.exists():
            print(f"Error: {md_path} not found", file=sys.stderr)
            sys.exit(1)
        md_text = md_path.read_text()
    else:
        parser.print_help()
        sys.exit(1)

    # Resolve style
    try:
        style = get_style(args.style, args.style_file)
    except (ValueError, FileNotFoundError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # Convert
    try:
        url = convert(
            md_text=md_text,
            doc_id=args.doc_id,
            title=args.title,
            style=style,
            token_path=args.credentials,
            embed=not args.no_diagrams,
            quiet=args.quiet,
        )
        if args.quiet:
            print(url)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
