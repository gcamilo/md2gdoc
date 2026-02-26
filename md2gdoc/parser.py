"""Markdown parser — converts markdown text to a list of typed blocks."""

import re


def _match_link(text, pos):
    """Match a markdown link at position pos, handling balanced parentheses in URLs.

    Returns (link_text, url, total_consumed_length) or None.
    """
    if pos >= len(text) or text[pos] != '[':
        return None

    # Find closing bracket for link text
    bracket_depth = 0
    j = pos
    while j < len(text):
        if text[j] == '[':
            bracket_depth += 1
        elif text[j] == ']':
            bracket_depth -= 1
            if bracket_depth == 0:
                break
        j += 1
    else:
        return None

    link_text = text[pos + 1:j]
    j += 1  # past ']'

    if j >= len(text) or text[j] != '(':
        return None
    j += 1  # past '('

    # Parse URL with balanced parentheses
    url_start = j
    paren_depth = 1
    while j < len(text) and paren_depth > 0:
        if text[j] == '(':
            paren_depth += 1
        elif text[j] == ')':
            paren_depth -= 1
        j += 1

    if paren_depth != 0:
        return None

    url = text[url_start:j - 1]  # j-1 because j is past the closing ')'
    total_len = j - pos
    return link_text, url, total_len


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
        # Link [text](url) — handles balanced parentheses in URL
        link = _match_link(text, j)
        if link:
            link_text, url, consumed = link
            s = pos
            result += link_text
            pos += len(link_text)
            formats.append((s, pos, "link", url))
            j += consumed
            continue
        result += text[j]
        pos += 1
        j += 1
    return result, formats


def parse_markdown(md_text: str) -> list[dict]:
    """Parse markdown into a list of blocks.

    Block types: heading, table, bullet, numbered, blockquote, paragraph,
    diagram, code_block.

    Consecutive non-blank, non-block lines are accumulated into a single
    paragraph block. Only blank lines separate paragraphs.
    """
    lines = md_text.split("\n")
    blocks = []
    i = 0
    para_accum = []  # accumulates consecutive paragraph lines

    def _flush_para():
        """Flush accumulated paragraph lines as a single paragraph block."""
        if para_accum:
            blocks.append({"type": "paragraph", "text": " ".join(para_accum)})
            para_accum.clear()

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Empty lines and horizontal rules flush paragraph and are skipped
        if not stripped or stripped in ("---", "- --", "***", "___"):
            _flush_para()
            i += 1
            continue

        # Fenced code block
        if stripped.startswith("```"):
            _flush_para()
            # Extract optional language hint
            lang = stripped[3:].strip() or None
            code_lines = []
            i += 1
            while i < len(lines):
                if lines[i].strip().startswith("```"):
                    i += 1
                    break
                code_lines.append(lines[i])
                i += 1
            blocks.append({
                "type": "code_block",
                "text": "\n".join(code_lines),
                "language": lang,
            })
            continue

        # Heading
        hm = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if hm:
            _flush_para()
            blocks.append({"type": "heading", "text": hm.group(2), "level": len(hm.group(1))})
            i += 1
            continue

        # Table
        if "|" in stripped and stripped.startswith("|"):
            _flush_para()
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
            _flush_para()
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
            _flush_para()
            blocks.append({"type": "blockquote", "text": stripped.lstrip("> ").strip()})
            i += 1
            continue

        # Bullet list
        bm = re.match(r"^(\s*)-\s+(.+)$", line)
        if bm:
            _flush_para()
            blocks.append({"type": "bullet", "text": bm.group(2), "level": len(bm.group(1)) // 2})
            i += 1
            continue

        # Numbered list
        nm = re.match(r"^(\s*)\d+\.\s+(.+)$", line)
        if nm:
            _flush_para()
            blocks.append({"type": "numbered", "text": nm.group(2), "level": len(nm.group(1)) // 2})
            i += 1
            continue

        # Regular text — accumulate for paragraph merging
        para_accum.append(stripped)
        i += 1

    # Flush any remaining accumulated paragraph lines
    _flush_para()

    return blocks
