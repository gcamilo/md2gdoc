---
name: md2gdoc
description: Convert markdown files to formatted Google Docs with style presets. No pandoc.
allowed-tools: Bash, Read, Glob
---

# md2gdoc — Markdown → Google Docs

Convert markdown to professionally formatted Google Docs via the Google Docs API.
No pandoc, no intermediate formats — direct API calls with pluggable style presets.

## When to Use

- Converting any `.md` file to a Google Doc
- Applying professional formatting (IB pitch-book, consulting, legal, etc.)
- Overwriting an existing Google Doc with updated markdown content
- Piping generated markdown directly to a formatted doc

## Prerequisites

- Google OAuth token at `~/.config/md2gdoc/token.json`
- If not authenticated: `python3 -m md2gdoc auth --client-secret /path/to/creds.json`
- Falls back to `~/.config/google-docs-mcp/token.json` if md2gdoc token doesn't exist

## Usage

```bash
# Basic — new doc, no styling
python3 -m md2gdoc input.md

# With style preset
python3 -m md2gdoc input.md --style ib
python3 -m md2gdoc input.md --style clean
python3 -m md2gdoc input.md --style consulting

# Overwrite existing doc
python3 -m md2gdoc input.md --doc-id DOC_ID --style ib

# Custom title
python3 -m md2gdoc input.md --style ib --title "Q4 Report"

# From stdin (useful for piping generated content)
echo "# Hello\n\nWorld" | python3 -m md2gdoc --stdin --style clean

# Custom style from JSON (deep-merged with 'clean' defaults)
python3 -m md2gdoc input.md --style-file my-style.json

# Quiet mode (only prints doc URL)
python3 -m md2gdoc input.md --style ib -q

# List all presets
python3 -m md2gdoc --list-styles

# Skip diagram embedding
python3 -m md2gdoc input.md --style ib --no-diagrams
```

## Style Presets (11)

| Preset | Description |
|--------|-------------|
| `ib` | Investment Banking — navy/gold, Georgia body, Arial tables |
| `clean` | Minimal — Arial throughout, subtle gray headings |
| `corporate` | Blue headers, professional table styling |
| `academic` | Times New Roman, wide margins, formal spacing |
| `memo` | Georgia body, compact, executive-style |
| `startup` | Inter font, teal accents, modern feel |
| `consulting` | Calibri, navy/orange, McKinsey-style |
| `legal` | Garamond, numbered headings, tight margins |
| `creative` | Playfair Display headings, Source Serif body |
| `report` | Cambria body, slate headers, formal |
| `minimal` | System fonts, near-zero chrome |
| `none` | No formatting — content insertion only |

## 4-Phase Pipeline

1. **Insert content** — text, tables, bullets, code blocks, diagram placeholders
2. **Embed diagrams** — export Google Slides as PNG, insert inline (private by default)
3. **Apply styling** — fonts, colors, heading sizes, table formatting, margins
4. **Readability spacing** — paragraph breathing room, bullet transitions, table padding

## Custom Style JSON

Partial JSON is deep-merged with `clean` preset defaults:

```json
{
  "body": {"font": "Helvetica", "size": 12},
  "headings": {"color": {"red": 0.2, "green": 0.0, "blue": 0.5}}
}
```

Only override what you need — everything else falls back to clean.

## Known Limitations

- No `![alt](url)` standard image syntax (only `[DIAGRAM: ...]` from Google Slides)
- No nested formatting (`**bold *italic***` — outer wins)
- No `_underscore_` italic or `__bold__` syntax (only `*` and `**`)
- No `~~strikethrough~~`
- Custom fonts (Inter, Playfair) may silently fall back to Arial if not added to user's Google Fonts
- `---` thematic breaks are skipped (not converted to horizontal rules)

## Package Location

`~/md2gdoc/` — standalone package, not inside personal repo.

## Troubleshooting

- **"No Google OAuth token found"** → Run `python3 -m md2gdoc auth --setup-guide`
- **Doc renders in Arial instead of custom font** → Manually add font via Google Docs > More Fonts
- **Rate limit errors** → Built-in retry with exponential backoff (up to 5 attempts)
- **`--style-file` crashes** → Ensure valid JSON; partial files are fine (deep-merged)
