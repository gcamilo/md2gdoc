"""Markdown parser — converts markdown text to a list of typed blocks."""

import re


def strip_md(text: str) -> tuple[str, list]:
    """Remove markdown inline formatting, return (plain_text, format_ranges).

    Format ranges are tuples of (start, end, type[, extra]).
    Types: "bold", "italic", "link" (with URL as extra), "code".
    """
    result = ""
    formats = []
    j = 0
    pos = 0
    while j < len(text):
        # Inline code `text`
        m = re.match(r"`([^`]+)`", text[j:])
        if m:
            s = pos
            result += m.group(1)
            pos += len(m.group(1))
            formats.append((s, pos, "code"))
            j += m.end()
            continue
        # Bold **text**
        m = re.match(r"\*\*(.+?)\*\*", text[j:])
        if m:
            s = pos
            result += m.group(1)
            pos += len(m.group(1))
            formats.append((s, pos, "bold"))
            j += m.end()
            continue
        # Italic *text*
        m = re.match(r"\*([^*]+?)\*", text[j:])
        if m:
            s = pos
            result += m.group(1)
            pos += len(m.group(1))
            formats.append((s, pos, "italic"))
            j += m.end()
            continue
        # Link [text](url)
        m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", text[j:])
        if m:
            s = pos
            result += m.group(1)
            pos += len(m.group(1))
            formats.append((s, pos, "link", m.group(2)))
            j += m.end()
            continue
        result += text[j]
        pos += 1
        j += 1
    return result, formats


def parse_markdown(md_text: str) -> list[dict]:
    """Parse markdown into a list of blocks.

    Block types: heading, table, bullet, numbered, blockquote, paragraph, diagram.
    """
    lines = md_text.split("\n")
    blocks = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Skip empty lines and horizontal rules
        if not stripped or stripped in ("---", "- --", "***", "___"):
            i += 1
            continue

        # Heading
        hm = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if hm:
            blocks.append({"type": "heading", "text": hm.group(2), "level": len(hm.group(1))})
            i += 1
            continue

        # Table
        if "|" in stripped and stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tl = lines[i].strip()
                # Skip separator rows (|---|---|)
                if not re.match(r"^\|[\s\-:|]+\|$", tl):
                    cells = [c.strip() for c in tl.split("|")[1:-1]]
                    table_lines.append(cells)
                i += 1
            if table_lines:
                blocks.append({"type": "table", "rows": table_lines})
            continue

        # Diagram reference: *[DIAGRAM: Label -- [Slide N](URL)]*
        dm = re.match(
            r"^\*?\[DIAGRAM:\s*(.+?)\s*[—\-]+\s*\[Slide\s+(\d+)\]\(https://docs\.google\.com/presentation/d/([^/]+)/[^)]*\)\]\*?$",
            stripped,
        )
        if dm:
            blocks.append({
                "type": "diagram",
                "label": dm.group(1),
                "slide_num": int(dm.group(2)),
                "presentation_id": dm.group(3),
            })
            i += 1
            continue

        # Blockquote
        if stripped.startswith(">"):
            blocks.append({"type": "blockquote", "text": stripped.lstrip("> ").strip()})
            i += 1
            continue

        # Bullet list
        bm = re.match(r"^(\s*)-\s+(.+)$", line)
        if bm:
            blocks.append({"type": "bullet", "text": bm.group(2), "level": len(bm.group(1)) // 2})
            i += 1
            continue

        # Numbered list
        nm = re.match(r"^(\s*)\d+\.\s+(.+)$", line)
        if nm:
            blocks.append({"type": "numbered", "text": nm.group(2), "level": len(nm.group(1)) // 2})
            i += 1
            continue

        # Regular paragraph
        blocks.append({"type": "paragraph", "text": stripped})
        i += 1

    return blocks
