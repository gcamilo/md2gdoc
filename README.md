# md2gdoc

Convert Markdown to professionally formatted Google Docs. No pandoc, no intermediate formats — direct Google Docs API calls with pluggable style presets.

## Install

```bash
pip install .
```

Or use directly:

```bash
python3 -m md2gdoc input.md --style clean
```

## Quick Start

```bash
# 1. Set up Google credentials (one-time, ~3 minutes)
md2gdoc auth --setup-guide          # Shows step-by-step instructions
md2gdoc auth --client-secret creds.json  # Authenticate (opens browser)

# 2. Convert your first doc
md2gdoc input.md --style clean

# 3. Try different styles
md2gdoc input.md --style ib          # Investment banking pitch-book
md2gdoc input.md --style consulting  # McKinsey-style
md2gdoc input.md --style academic    # Times New Roman, formal
```

## Style Presets

| Preset | Look |
|--------|------|
| `ib` | Investment banking — navy/gold, Georgia body, Arial tables |
| `clean` | Minimal — Arial throughout, subtle gray headings |
| `corporate` | Blue headers, professional table styling |
| `academic` | Times New Roman, wide margins, formal spacing |
| `memo` | Georgia body, compact executive-style |
| `startup` | Inter font, teal accents, modern |
| `consulting` | Calibri, navy/orange, McKinsey-style |
| `legal` | Garamond, numbered headings, tight margins |
| `creative` | Playfair Display headings, Source Serif body |
| `report` | Cambria body, slate headers, formal |
| `minimal` | System fonts, near-zero chrome |

Use `--style none` (or omit `--style`) to insert content without formatting.

## Custom Styles

Pass a JSON file with `--style-file`. Partial definitions are deep-merged with the `clean` preset, so you only need to override what you want:

```json
{
  "body": { "font": "Helvetica", "size": 12 },
  "headings": { "color": { "red": 0.2, "green": 0.0, "blue": 0.5 } }
}
```

```bash
md2gdoc input.md --style-file my-brand.json
```

## CLI Reference

```
md2gdoc input.md                         # New doc, no styling
md2gdoc input.md --style ib              # With style preset
md2gdoc input.md --doc-id DOC_ID         # Overwrite existing doc
md2gdoc input.md --title "Q4 Report"     # Custom title
md2gdoc input.md --style-file brand.json # Custom style from JSON
md2gdoc --stdin --style clean            # Read from stdin
md2gdoc --list-styles                    # Show all presets
md2gdoc input.md -q                      # Quiet (prints only URL)
md2gdoc input.md --no-diagrams           # Skip diagram embedding
md2gdoc auth --check                     # Verify credentials
md2gdoc auth --revoke                    # Delete stored credentials
```

## How It Works

1. **Parse** — Markdown → typed blocks (headings, paragraphs, tables, bullets, code blocks, blockquotes)
2. **Insert** — Blocks → Google Docs via `batchUpdate` API calls
3. **Style** — Apply fonts, colors, heading sizes, table formatting, margins
4. **Space** — Paragraph breathing room, bullet transitions, table padding

Diagrams from Google Slides are optionally embedded as inline images (requires Slides API access).

## Supported Markdown

- Headings (`# H1` through `###### H6`)
- Paragraphs (consecutive lines merged)
- Bold (`**text**`), italic (`*text*`), inline code (`` `text` ``)
- Links (`[text](url)`) with balanced parentheses
- Bullet lists (`- item`, nested via indentation)
- Numbered lists (`1. item`, nested via indentation)
- Tables (pipe-delimited with header separator)
- Blockquotes (`> text`)
- Fenced code blocks (`` ```language ``)
- Diagram references (`[DIAGRAM: Label -- [Slide N](slides-url)]`)

## Credentials Setup

md2gdoc needs OAuth2 credentials to create Google Docs on your behalf.

1. Go to [Google Cloud Console](https://console.cloud.google.com/apis/credentials)
2. Create a project and enable **Google Docs API** and **Google Drive API**
3. Create an **OAuth client ID** (Desktop app type)
4. Download the JSON file
5. Run: `md2gdoc auth --client-secret /path/to/downloaded.json`

This opens a browser for authorization. Credentials are stored at `~/.config/md2gdoc/token.json` (chmod 600). You only need to do this once.

## Limitations

- No `![alt](url)` standard image embedding (only Google Slides diagrams)
- No nested formatting (`**bold *and italic***` — outer wins)
- No underscore syntax (`_italic_`, `__bold__`)
- No strikethrough (`~~text~~`)
- Custom Google Fonts (Inter, Playfair Display) may fall back to Arial if not manually added via Google Docs → More Fonts
- Thematic breaks (`---`) are treated as paragraph separators, not horizontal rules

## License

MIT
