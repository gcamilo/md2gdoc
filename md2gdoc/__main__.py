"""CLI entry point for md2gdoc."""

import argparse
import sys
from pathlib import Path

from . import __version__
from .styles import get_style, list_styles


def cmd_auth(args):
    """Handle auth subcommand."""
    from .auth import run_auth_flow, check_auth, revoke_auth, SETUP_GUIDE

    if args.revoke:
        revoke_auth()
    elif args.check:
        if check_auth():
            print("Authenticated. Credentials are valid.")
        else:
            print("Not authenticated. Run: md2gdoc auth --client-secret /path/to/secret.json")
            sys.exit(1)
    elif args.setup_guide:
        print(SETUP_GUIDE)
    else:
        run_auth_flow(args.client_secret)


def cmd_convert(args):
    """Handle convert (default) subcommand."""
    from .engine import convert

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
        print("Error: provide a markdown file or use --stdin", file=sys.stderr)
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
            public_diagrams=getattr(args, "public_diagrams", False),
        )
        if args.quiet:
            print(url)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        print("\nRun 'md2gdoc auth' to set up credentials.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        prog="md2gdoc",
        description="Convert Markdown to formatted Google Docs. No pandoc.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  md2gdoc auth --client-secret creds.json   Set up Google credentials (once)
  md2gdoc input.md                           Create new doc (no styling)
  md2gdoc input.md --style ib                IB pitch-book formatting
  md2gdoc input.md --style clean             Clean, minimal formatting
  md2gdoc input.md --style corporate         Corporate blue theme
  md2gdoc input.md --doc-id DOC_ID           Overwrite existing doc
  md2gdoc input.md --style ib --title "Q4"   Custom title + IB style
  md2gdoc input.md --style-file my.json      Custom style from JSON
  md2gdoc --list-styles                      Show available style presets
  md2gdoc --stdin --style ib                 Read markdown from stdin

getting started:
  1. md2gdoc auth --setup-guide              Show credential setup instructions
  2. md2gdoc auth --client-secret creds.json Authenticate (opens browser)
  3. md2gdoc input.md --style clean          Convert your first doc
        """,
    )

    subparsers = parser.add_subparsers(dest="command")

    # Auth subcommand
    auth_parser = subparsers.add_parser("auth", help="Set up Google credentials")
    auth_parser.add_argument("--client-secret",
                             help="Path to OAuth2 client secret JSON from Google Cloud Console")
    auth_parser.add_argument("--check", action="store_true",
                             help="Check if credentials are valid")
    auth_parser.add_argument("--revoke", action="store_true",
                             help="Revoke and delete stored credentials")
    auth_parser.add_argument("--setup-guide", action="store_true",
                             help="Show step-by-step credential setup guide")

    # Convert arguments (on the main parser for backwards compat)
    parser.add_argument("input", nargs="?", help="Markdown file path")
    parser.add_argument("--doc-id", help="Existing Google Doc ID to overwrite")
    parser.add_argument("--title", help="Document title (defaults to first H1)")
    parser.add_argument("--style", default=None,
                        help="Style preset: ib, clean, corporate, none (default: none)")
    parser.add_argument("--style-file", help="Path to custom style JSON file")
    parser.add_argument("--credentials", help="Path to Google OAuth token JSON")
    parser.add_argument("--no-diagrams", action="store_true",
                        help="Skip diagram embedding from Google Slides")
    parser.add_argument("--public-diagrams", action="store_true",
                        help="Make uploaded diagram images world-readable (default: private)")
    parser.add_argument("--stdin", action="store_true",
                        help="Read markdown from stdin instead of file")
    parser.add_argument("--quiet", "-q", action="store_true",
                        help="Suppress progress output")
    parser.add_argument("--list-styles", action="store_true",
                        help="List available style presets and exit")
    parser.add_argument("--version", "-v", action="version",
                        version=f"md2gdoc {__version__}")

    args = parser.parse_args()

    # Handle --list-styles
    if getattr(args, "list_styles", False):
        print("Available style presets:\n")
        for name, desc in list_styles():
            print(f"  {name:12s}  {desc}")
        print(f"\n  {'none':12s}  No formatting (content only)")
        print("\nUse --style-file for custom styles (JSON).")
        sys.exit(0)

    # Route to subcommand
    if args.command == "auth":
        cmd_auth(args)
    elif args.input or getattr(args, "stdin", False):
        cmd_convert(args)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
