"""Google Docs API engine — inserts content and applies styling."""

import json
import os
import random
import re
import tempfile
import time
import urllib.request
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from .parser import strip_md

PAGE_WIDTH_PT = 468  # 8.5in - 1in margins each side


# ── Auth ───────────────────────────────────────────────────────────────

def get_credentials(token_path: str | None = None) -> Credentials:
    """Load OAuth2 credentials from a token JSON file.

    Searches in order:
    1. Explicit token_path argument
    2. GOOGLE_TOKEN_PATH environment variable
    3. ~/.config/md2gdoc/token.json
    4. ~/.config/google-docs-mcp/token.json (compatibility)
    """
    import os

    candidates = []
    if token_path:
        candidates.append(Path(token_path))
    env_path = os.environ.get("GOOGLE_TOKEN_PATH")
    if env_path:
        candidates.append(Path(env_path))
    candidates.append(Path.home() / ".config/md2gdoc/token.json")
    candidates.append(Path.home() / ".config/google-docs-mcp/token.json")

    for p in candidates:
        if p.exists():
            with open(p) as f:
                td = json.load(f)
            creds = Credentials(
                token=td.get("access_token") or td.get("token"),
                refresh_token=td.get("refresh_token"),
                token_uri="https://oauth2.googleapis.com/token",
                client_id=td.get("client_id"),
                client_secret=td.get("client_secret"),
            )
            try:
                creds.refresh(Request())
            except Exception as refresh_err:
                # If refresh fails, check whether the existing token is still valid
                if creds.token and creds.valid:
                    return creds
                # No refresh_token means we can't recover — guide the user
                if not creds.refresh_token:
                    raise RuntimeError(
                        "Credential refresh failed and no refresh_token is available.\n"
                        "Run 'md2gdoc auth' to re-authenticate."
                    ) from refresh_err
                raise
            return creds

    searched = [str(p) for p in candidates]
    raise FileNotFoundError(
        f"No Google OAuth token found. Searched:\n"
        + "\n".join(f"  - {p}" for p in searched)
        + "\n\nRun 'md2gdoc auth' to set up credentials."
        + "\nOr run 'md2gdoc auth --setup-guide' for step-by-step instructions."
    )


def get_docs_service(creds: Credentials = None, token_path: str = None):
    if creds is None:
        creds = get_credentials(token_path)
    return build("docs", "v1", credentials=creds)


def get_slides_service(creds: Credentials = None, token_path: str = None):
    if creds is None:
        creds = get_credentials(token_path)
    return build("slides", "v1", credentials=creds)


def get_drive_service(creds: Credentials = None, token_path: str = None):
    if creds is None:
        creds = get_credentials(token_path)
    return build("drive", "v3", credentials=creds)


# ── Helpers ────────────────────────────────────────────────────────────

def _is_retryable(exc: Exception) -> bool:
    """Return True if the exception represents a retryable API error (429 / 5xx)."""
    if isinstance(exc, HttpError):
        status = exc.resp.status if hasattr(exc, "resp") else 0
        return status == 429 or 500 <= status < 600
    # Fallback: string matching for wrapped errors
    msg = str(exc)
    return "429" in msg or "RATE_LIMIT" in msg or "500" in msg or "503" in msg


def batch_send(docs_svc, doc_id: str, reqs: list, label: str = "",
               batch_size: int = 80, quiet: bool = False):
    """Send requests in batches with rate-limit retry and exponential backoff + jitter."""
    if not reqs:
        return
    for i in range(0, len(reqs), batch_size):
        for attempt in range(5):
            try:
                docs_svc.documents().batchUpdate(
                    documentId=doc_id,
                    body={"requests": reqs[i : i + batch_size]},
                ).execute()
                break
            except Exception as e:
                if _is_retryable(e) and attempt < 4:
                    base_wait = 15 * (attempt + 1)
                    jitter = random.uniform(0, base_wait * 0.3)
                    wait = base_wait + jitter
                    if not quiet:
                        print(f"  Retryable error, waiting {wait:.1f}s (attempt {attempt + 1}/5)...")
                    time.sleep(wait)
                else:
                    raise
    if label and not quiet:
        print(f"  {label}: {len(reqs)} reqs")


def get_doc_end(docs_svc, doc_id: str) -> int:
    doc = docs_svc.documents().get(documentId=doc_id).execute()
    return doc["body"]["content"][-1]["endIndex"]


def get_doc(docs_svc, doc_id: str) -> dict:
    return docs_svc.documents().get(documentId=doc_id).execute()


# ── Phase 1: Insert Content ───────────────────────────────────────────

def insert_content(docs_svc, doc_id: str, blocks: list, style: dict | None = None,
                   quiet: bool = False):
    """Insert all text/table content into the doc."""
    # Group blocks into segments (separated by tables and diagrams)
    segments = []
    current_seg = []
    for b in blocks:
        if b["type"] in ("table", "diagram"):
            if current_seg:
                segments.append({"type": "text", "blocks": current_seg})
                current_seg = []
            segments.append({"type": b["type"], "data": b})
        else:
            current_seg.append(b)
    if current_seg:
        segments.append({"type": "text", "blocks": current_seg})

    n_tables = sum(1 for s in segments if s["type"] == "table")
    if not quiet:
        print(f"  {len(segments)} segments ({n_tables} tables)")

    bq_style = (style or {}).get("blockquote", {})
    bq_color = bq_style.get("border_color", {"red": 0.5, "green": 0.5, "blue": 0.5})
    bq_width = bq_style.get("border_width", 2)
    bq_indent = bq_style.get("indent", 18)

    for seg_idx, seg in enumerate(segments):
        if seg["type"] == "text":
            cursor = get_doc_end(docs_svc, doc_id) - 1
            reqs = []
            fmt_reqs = []
            lc = cursor

            for b in seg["blocks"]:
                if b["type"] == "code_block":
                    # Code blocks: use raw text, no markdown processing
                    plain = b["text"]
                    fmts = []
                else:
                    plain, fmts = strip_md(b["text"])
                text = plain + "\n"
                reqs.append({"insertText": {"location": {"index": lc}, "text": text}})

                if b["type"] == "heading":
                    level = min(b["level"], 6)
                    reqs.append({"updateParagraphStyle": {
                        "range": {"startIndex": lc, "endIndex": lc + len(text)},
                        "paragraphStyle": {"namedStyleType": f"HEADING_{level}"},
                        "fields": "namedStyleType",
                    }})
                elif b["type"] == "bullet":
                    reqs.append({"createParagraphBullets": {
                        "range": {"startIndex": lc, "endIndex": lc + len(text)},
                        "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE",
                    }})
                    level = b.get("level", 0)
                    if level > 0:
                        indent_pt = 36 * level
                        reqs.append({"updateParagraphStyle": {
                            "range": {"startIndex": lc, "endIndex": lc + len(text)},
                            "paragraphStyle": {
                                "indentStart": {"magnitude": indent_pt, "unit": "PT"},
                                "indentFirstLine": {"magnitude": indent_pt, "unit": "PT"},
                            },
                            "fields": "indentStart,indentFirstLine",
                        }})
                elif b["type"] == "numbered":
                    reqs.append({"createParagraphBullets": {
                        "range": {"startIndex": lc, "endIndex": lc + len(text)},
                        "bulletPreset": "NUMBERED_DECIMAL_ALPHA_ROMAN",
                    }})
                    level = b.get("level", 0)
                    if level > 0:
                        indent_pt = 36 * level
                        reqs.append({"updateParagraphStyle": {
                            "range": {"startIndex": lc, "endIndex": lc + len(text)},
                            "paragraphStyle": {
                                "indentStart": {"magnitude": indent_pt, "unit": "PT"},
                                "indentFirstLine": {"magnitude": indent_pt, "unit": "PT"},
                            },
                            "fields": "indentStart,indentFirstLine",
                        }})
                elif b["type"] == "blockquote" and style:
                    fmt_reqs.append({"updateParagraphStyle": {
                        "range": {"startIndex": lc, "endIndex": lc + len(text)},
                        "paragraphStyle": {
                            "indentStart": {"magnitude": bq_indent, "unit": "PT"},
                            "indentEnd": {"magnitude": bq_indent, "unit": "PT"},
                            "borderLeft": {
                                "color": {"color": {"rgbColor": bq_color}},
                                "width": {"magnitude": bq_width, "unit": "PT"},
                                "padding": {"magnitude": 8, "unit": "PT"},
                                "dashStyle": "SOLID",
                            },
                        },
                        "fields": "indentStart,indentEnd,borderLeft",
                    }})
                elif b["type"] == "code_block":
                    # Monospace font + light gray background for code blocks
                    fmt_reqs.append({"updateTextStyle": {
                        "range": {"startIndex": lc, "endIndex": lc + len(text)},
                        "textStyle": {
                            "weightedFontFamily": {"fontFamily": "Courier New", "weight": 400},
                            "fontSize": {"magnitude": 10, "unit": "PT"},
                            "backgroundColor": {"color": {"rgbColor": {
                                "red": 0.95, "green": 0.95, "blue": 0.95}}},
                        },
                        "fields": "weightedFontFamily,fontSize,backgroundColor",
                    }})
                    fmt_reqs.append({"updateParagraphStyle": {
                        "range": {"startIndex": lc, "endIndex": lc + len(text)},
                        "paragraphStyle": {
                            "indentStart": {"magnitude": 12, "unit": "PT"},
                            "indentEnd": {"magnitude": 12, "unit": "PT"},
                            "spaceAbove": {"magnitude": 6, "unit": "PT"},
                            "spaceBelow": {"magnitude": 6, "unit": "PT"},
                        },
                        "fields": "indentStart,indentEnd,spaceAbove,spaceBelow",
                    }})

                # Inline formatting
                for fmt in fmts:
                    ts = {}
                    fields = []
                    if fmt[2] == "bold":
                        ts["bold"] = True
                        fields.append("bold")
                    elif fmt[2] == "italic":
                        ts["italic"] = True
                        fields.append("italic")
                    elif fmt[2] == "code":
                        ts["weightedFontFamily"] = {"fontFamily": "Courier New", "weight": 400}
                        ts["backgroundColor"] = {"color": {"rgbColor": {"red": 0.95, "green": 0.95, "blue": 0.95}}}
                        fields.extend(["weightedFontFamily", "backgroundColor"])
                    elif fmt[2] == "link":
                        ts["link"] = {"url": fmt[3]}
                        fields.append("link")
                    if ts:
                        fmt_reqs.append({"updateTextStyle": {
                            "range": {"startIndex": lc + fmt[0], "endIndex": lc + fmt[1]},
                            "textStyle": ts,
                            "fields": ",".join(fields),
                        }})

                lc += len(text)

            batch_send(docs_svc, doc_id, reqs,
                       f"Seg {seg_idx} ({len(seg['blocks'])} blocks)", quiet=quiet)
            batch_send(docs_svc, doc_id, fmt_reqs, "", quiet=True)

        elif seg["type"] == "table":
            rows = seg["data"]["rows"]
            num_rows = len(rows)
            num_cols = max(len(r) for r in rows)

            cursor = get_doc_end(docs_svc, doc_id) - 1
            docs_svc.documents().batchUpdate(documentId=doc_id, body={"requests": [
                {"insertTable": {"rows": num_rows, "columns": num_cols,
                                 "location": {"index": cursor}}}
            ]}).execute()

            # Re-read to get cell indices
            doc = get_doc(docs_svc, doc_id)
            tables_in_doc = [el for el in doc["body"]["content"] if "table" in el]
            t = tables_in_doc[-1]["table"]

            # Populate cells in REVERSE order to avoid index shifting
            cell_reqs = []
            for row_idx in range(num_rows - 1, -1, -1):
                t_rows = t["tableRows"]
                if row_idx >= len(t_rows):
                    continue
                for col_idx in range(num_cols - 1, -1, -1):
                    if col_idx >= len(t_rows[row_idx]["tableCells"]):
                        continue
                    cell = t_rows[row_idx]["tableCells"][col_idx]
                    content = cell.get("content", [])
                    if not content:
                        continue
                    cell_start = content[0].get("startIndex")
                    if cell_start is None:
                        continue
                    cell_text = rows[row_idx][col_idx] if col_idx < len(rows[row_idx]) else ""
                    plain, _fmts = strip_md(cell_text)
                    if plain:
                        cell_reqs.append({"insertText": {
                            "location": {"index": cell_start}, "text": plain}})

            batch_send(docs_svc, doc_id, cell_reqs,
                       f"Seg {seg_idx} table ({num_rows}x{num_cols})", quiet=quiet)

        elif seg["type"] == "diagram":
            cursor = get_doc_end(docs_svc, doc_id) - 1
            d = seg["data"]
            marker = f"{{{{DIAGRAM:{d['presentation_id']}:SLIDE{d['slide_num']}:{d['label']}}}}}\n"
            docs_svc.documents().batchUpdate(documentId=doc_id, body={"requests": [
                {"insertText": {"location": {"index": cursor}, "text": marker}}
            ]}).execute()
            if not quiet:
                print(f"  Diagram placeholder: {d['label']} (Slide {d['slide_num']})")


# ── Phase 2: Apply Styling ─────────────────────────────────────────────

def apply_styling(docs_svc, doc_id: str, style: dict, quiet: bool = False):
    """Apply a style preset to the entire document."""
    if not style:
        return

    body_cfg = style["body"]
    head_cfg = style["headings"]
    tbl_cfg = style["tables"]
    margins = style.get("margins", {})

    # Base font
    if not quiet:
        print("  Base font...")
    doc = get_doc(docs_svc, doc_id)
    doc_end = doc["body"]["content"][-1]["endIndex"]
    batch_send(docs_svc, doc_id, [{"updateTextStyle": {
        "range": {"startIndex": 1, "endIndex": doc_end - 1},
        "textStyle": {
            "weightedFontFamily": {"fontFamily": body_cfg["font"], "weight": 400},
            "fontSize": {"magnitude": body_cfg["size"], "unit": "PT"},
            "foregroundColor": {"color": {"rgbColor": body_cfg["color"]}},
        },
        "fields": "weightedFontFamily,fontSize,foregroundColor",
    }}], quiet=quiet)

    # Headings
    if not quiet:
        print("  Headings...")
    doc = get_doc(docs_svc, doc_id)
    body = doc["body"]["content"]
    reqs = []
    for el in body:
        if "paragraph" not in el:
            continue
        p = el["paragraph"]
        named = p.get("paragraphStyle", {}).get("namedStyleType", "")
        if not named.startswith("HEADING"):
            continue
        elements = p.get("elements", [])
        if not elements:
            continue
        ps = elements[0].get("startIndex")
        pe = elements[-1].get("endIndex")
        if ps is None or pe is None or pe <= ps:
            continue

        level = int(named.split("_")[1]) if "_" in named else 1
        reqs.append({"updateParagraphStyle": {
            "range": {"startIndex": ps, "endIndex": pe},
            "paragraphStyle": {
                "spaceAbove": {"magnitude": head_cfg["space_above"].get(level, 14), "unit": "PT"},
                "spaceBelow": {"magnitude": head_cfg["space_below"].get(level, 6), "unit": "PT"},
                "keepWithNext": True,
            },
            "fields": "spaceAbove,spaceBelow,keepWithNext",
        }})

        # H2 border
        h2_border = style.get("h2_border")
        if named == "HEADING_2" and h2_border:
            reqs.append({"updateParagraphStyle": {
                "range": {"startIndex": ps, "endIndex": pe},
                "paragraphStyle": {"borderBottom": {
                    "color": {"color": {"rgbColor": h2_border}},
                    "width": {"magnitude": 1, "unit": "PT"},
                    "padding": {"magnitude": 3, "unit": "PT"},
                    "dashStyle": "SOLID",
                }},
                "fields": "borderBottom",
            }})

        for elem in elements:
            s, e = elem.get("startIndex"), elem.get("endIndex")
            if s is None or e is None or e <= s:
                continue
            reqs.append({"updateTextStyle": {
                "range": {"startIndex": s, "endIndex": e},
                "textStyle": {
                    "bold": head_cfg["bold"].get(level, False),
                    "fontSize": {"magnitude": head_cfg["sizes"].get(level, 11), "unit": "PT"},
                    "foregroundColor": {"color": {"rgbColor": head_cfg["color"]}},
                    "weightedFontFamily": {"fontFamily": head_cfg["font"], "weight": 400},
                },
                "fields": "bold,fontSize,foregroundColor,weightedFontFamily",
            }})
    batch_send(docs_svc, doc_id, reqs, "Headings", quiet=quiet)

    # Tables
    if not quiet:
        print("  Tables...")
    doc = get_doc(docs_svc, doc_id)
    body = doc["body"]["content"]
    tables = [el for el in body if "table" in el]

    for t_idx, table_el in enumerate(tables):
        table_start = table_el["startIndex"]
        t = table_el["table"]
        t_rows = t["tableRows"]
        num_cols = t["columns"]
        reqs = []

        # Column width optimization
        col_max_chars = [0] * num_cols
        for row in t_rows:
            for col_idx, cell in enumerate(row["tableCells"]):
                txt = ""
                for c in cell.get("content", []):
                    if "paragraph" in c:
                        for elem in c["paragraph"].get("elements", []):
                            txt += elem.get("textRun", {}).get("content", "")
                col_max_chars[col_idx] = max(col_max_chars[col_idx], len(txt.strip()))

        CHAR_W = 5.5
        MIN_COL = 45
        raw = []
        for chars in col_max_chars:
            if chars <= 8:
                raw.append(max(MIN_COL, chars * CHAR_W + 12))
            elif chars <= 20:
                raw.append(max(60, chars * CHAR_W + 12))
            elif chars <= 40:
                raw.append(max(80, min(chars * CHAR_W + 12, 160)))
            else:
                raw.append(max(120, min(chars * CHAR_W * 0.6, 300)))
        total = sum(raw)
        widths = [max(MIN_COL, w * PAGE_WIDTH_PT / total) for w in raw]
        wt = sum(widths)
        widths = [w * PAGE_WIDTH_PT / wt for w in widths]

        for col_idx in range(num_cols):
            reqs.append({"updateTableColumnProperties": {
                "tableStartLocation": {"index": table_start},
                "columnIndices": [col_idx],
                "tableColumnProperties": {
                    "widthType": "FIXED_WIDTH",
                    "width": {"magnitude": widths[col_idx], "unit": "PT"},
                },
                "fields": "widthType,width",
            }})

        # Cell styling
        for row_idx, row in enumerate(t_rows):
            is_header = row_idx == 0
            bg = (tbl_cfg["header_bg"] if is_header
                  else (tbl_cfg["alt_row_bg"] if row_idx % 2 == 1
                        else tbl_cfg["row_bg"]))
            padding = tbl_cfg["header_padding"] if is_header else tbl_cfg["cell_padding"]

            for col_idx, cell in enumerate(row["tableCells"]):
                reqs.append({"updateTableCellStyle": {
                    "tableRange": {
                        "tableCellLocation": {
                            "tableStartLocation": {"index": table_start},
                            "rowIndex": row_idx, "columnIndex": col_idx,
                        },
                        "rowSpan": 1, "columnSpan": 1,
                    },
                    "tableCellStyle": {
                        "backgroundColor": {"color": {"rgbColor": bg}},
                        "paddingTop": {"magnitude": padding["top"], "unit": "PT"},
                        "paddingBottom": {"magnitude": padding["bottom"], "unit": "PT"},
                        "paddingLeft": {"magnitude": padding["left"], "unit": "PT"},
                        "paddingRight": {"magnitude": padding["right"], "unit": "PT"},
                    },
                    "fields": "backgroundColor,paddingTop,paddingBottom,paddingLeft,paddingRight",
                }})

                content = cell.get("content", [])
                if not content:
                    continue
                s = content[0].get("startIndex")
                e = content[-1].get("endIndex")
                if s is None or e is None or e <= s:
                    continue

                text_color = tbl_cfg["header_color"] if is_header else tbl_cfg["body_color"]
                reqs.append({"updateTextStyle": {
                    "range": {"startIndex": s, "endIndex": e},
                    "textStyle": {
                        "bold": tbl_cfg["header_bold"] if is_header else False,
                        "foregroundColor": {"color": {"rgbColor": text_color}},
                        "fontSize": {"magnitude": tbl_cfg["size"], "unit": "PT"},
                        "weightedFontFamily": {"fontFamily": tbl_cfg["font"], "weight": 400},
                    },
                    "fields": "bold,foregroundColor,fontSize,weightedFontFamily",
                }})

                alignment = "CENTER" if col_idx > 0 else "START"
                reqs.append({"updateParagraphStyle": {
                    "range": {"startIndex": s, "endIndex": e},
                    "paragraphStyle": {
                        "alignment": alignment,
                        "spaceAbove": {"magnitude": 0, "unit": "PT"},
                        "spaceBelow": {"magnitude": 0, "unit": "PT"},
                        "lineSpacing": 100,
                    },
                    "fields": "alignment,spaceAbove,spaceBelow,lineSpacing",
                }})

        batch_send(docs_svc, doc_id, reqs,
                   f"Table {t_idx} ({len(t_rows)}r x {num_cols}c)",
                   batch_size=100, quiet=quiet)

    # Page margins
    if margins:
        if not quiet:
            print("  Margins...")
        batch_send(docs_svc, doc_id, [{"updateDocumentStyle": {
            "documentStyle": {
                "marginTop": {"magnitude": margins.get("top", 72), "unit": "PT"},
                "marginBottom": {"magnitude": margins.get("bottom", 72), "unit": "PT"},
                "marginLeft": {"magnitude": margins.get("left", 72), "unit": "PT"},
                "marginRight": {"magnitude": margins.get("right", 72), "unit": "PT"},
            },
            "fields": "marginTop,marginBottom,marginLeft,marginRight",
        }}], quiet=quiet)


# ── Phase 3: Readability Spacing ───────────────────────────────────────

def apply_spacing(docs_svc, doc_id: str, style: dict, quiet: bool = False):
    """Add breathing room: body text, bullets, tables, transitions."""
    if not style:
        return

    body_cfg = style["body"]
    bullet_cfg = style.get("bullets", {})

    # Collect table paragraph indices to skip
    doc = get_doc(docs_svc, doc_id)
    body = doc["body"]["content"]
    table_indices = set()
    for el in body:
        if "table" in el:
            for row in el["table"]["tableRows"]:
                for cell in row["tableCells"]:
                    for c in cell.get("content", []):
                        if "paragraph" in c:
                            for elem in c["paragraph"].get("elements", []):
                                si = elem.get("startIndex")
                                if si is not None:
                                    table_indices.add(si)

    def _para_range(p):
        elements = p.get("elements", [])
        if not elements:
            return None, None
        s = elements[0].get("startIndex")
        e = elements[-1].get("endIndex")
        if s is None or e is None or e <= s:
            return None, None
        return s, e

    def _para_text(p):
        t = ""
        for elem in p.get("elements", []):
            t += elem.get("textRun", {}).get("content", "")
        return t.strip()

    # Body paragraph spacing
    if not quiet:
        print("  Body paragraphs...")
    reqs = []
    for el in body:
        if "paragraph" not in el:
            continue
        p = el["paragraph"]
        named = p.get("paragraphStyle", {}).get("namedStyleType", "")
        if named.startswith("HEADING"):
            continue
        s, e = _para_range(p)
        if s is None or s in table_indices:
            continue

        is_bullet = "bullet" in p
        text = _para_text(p)

        if is_bullet:
            reqs.append({"updateParagraphStyle": {
                "range": {"startIndex": s, "endIndex": e},
                "paragraphStyle": {
                    "spaceAbove": {"magnitude": bullet_cfg.get("space_above", 1), "unit": "PT"},
                    "spaceBelow": {"magnitude": bullet_cfg.get("space_below", 3), "unit": "PT"},
                    "lineSpacing": bullet_cfg.get("line_spacing", 115),
                },
                "fields": "spaceAbove,spaceBelow,lineSpacing",
            }})
        elif text == "":
            reqs.append({"updateParagraphStyle": {
                "range": {"startIndex": s, "endIndex": e},
                "paragraphStyle": {
                    "spaceAbove": {"magnitude": 0, "unit": "PT"},
                    "spaceBelow": {"magnitude": 0, "unit": "PT"},
                    "lineSpacing": 50,
                },
                "fields": "spaceAbove,spaceBelow,lineSpacing",
            }})
        else:
            reqs.append({"updateParagraphStyle": {
                "range": {"startIndex": s, "endIndex": e},
                "paragraphStyle": {
                    "spaceAbove": {"magnitude": body_cfg.get("space_above", 2), "unit": "PT"},
                    "spaceBelow": {"magnitude": body_cfg.get("space_below", 8), "unit": "PT"},
                    "lineSpacing": body_cfg.get("line_spacing", 120),
                },
                "fields": "spaceAbove,spaceBelow,lineSpacing",
            }})
    batch_send(docs_svc, doc_id, reqs, "Body spacing", quiet=quiet)

    # Table padding
    if not quiet:
        print("  Table padding...")
    reqs = []
    for i, el in enumerate(body):
        if "table" not in el:
            continue
        if i > 0 and "paragraph" in body[i - 1]:
            s, e = _para_range(body[i - 1]["paragraph"])
            if s is not None:
                reqs.append({"updateParagraphStyle": {
                    "range": {"startIndex": s, "endIndex": e},
                    "paragraphStyle": {"spaceBelow": {"magnitude": 10, "unit": "PT"}},
                    "fields": "spaceBelow",
                }})
        if i + 1 < len(body) and "paragraph" in body[i + 1]:
            s, e = _para_range(body[i + 1]["paragraph"])
            if s is not None:
                reqs.append({"updateParagraphStyle": {
                    "range": {"startIndex": s, "endIndex": e},
                    "paragraphStyle": {"spaceAbove": {"magnitude": 10, "unit": "PT"}},
                    "fields": "spaceAbove",
                }})
    batch_send(docs_svc, doc_id, reqs, "Table padding", quiet=quiet)

    # Bullet transitions
    if not quiet:
        print("  Bullet transitions...")
    reqs = []
    prev_bullet = False
    for el in body:
        if "paragraph" not in el:
            prev_bullet = False
            continue
        p = el["paragraph"]
        is_bullet = "bullet" in p
        s, e = _para_range(p)
        if s is None or s in table_indices:
            prev_bullet = False
            continue
        if is_bullet and not prev_bullet:
            reqs.append({"updateParagraphStyle": {
                "range": {"startIndex": s, "endIndex": e},
                "paragraphStyle": {"spaceAbove": {
                    "magnitude": bullet_cfg.get("list_start_space", 6), "unit": "PT"}},
                "fields": "spaceAbove",
            }})
        if not is_bullet and prev_bullet:
            reqs.append({"updateParagraphStyle": {
                "range": {"startIndex": s, "endIndex": e},
                "paragraphStyle": {"spaceAbove": {
                    "magnitude": bullet_cfg.get("list_end_space", 8), "unit": "PT"}},
                "fields": "spaceAbove",
            }})
        prev_bullet = is_bullet
    batch_send(docs_svc, doc_id, reqs, "Bullet transitions", quiet=quiet)


# ── Phase 4: Embed Diagrams ───────────────────────────────────────────

def embed_diagrams(docs_svc, doc_id: str, blocks: list,
                   creds: Credentials = None, quiet: bool = False,
                   public_diagrams: bool = False):
    """Find diagram placeholders, export slides as PNG, insert inline images."""
    diagrams = [b for b in blocks if b["type"] == "diagram"]
    if not diagrams:
        return

    if not quiet:
        print(f"\n  Found {len(diagrams)} diagram(s) to embed")

    slides_svc = get_slides_service(creds)
    drive_svc = get_drive_service(creds)

    # Group by presentation
    pres_slides = {}
    for d in diagrams:
        pid = d["presentation_id"]
        if pid not in pres_slides:
            try:
                pres = slides_svc.presentations().get(presentationId=pid).execute()
                slide_map = {}
                for i, slide in enumerate(pres.get("slides", []), 1):
                    slide_map[i] = slide["objectId"]
                pres_slides[pid] = slide_map
                if not quiet:
                    print(f"  Loaded presentation {pid}: {len(slide_map)} slides")
            except Exception as e:
                print(f"  WARNING: Failed to load presentation {pid}: {e}")
                continue

    # Export each slide as PNG using secure temp files
    exported = {}  # key -> local_path
    temp_files = []  # track for cleanup
    for d in diagrams:
        key = (d["presentation_id"], d["slide_num"])
        if key in exported:
            continue
        pid, snum = key
        page_id = pres_slides.get(pid, {}).get(snum)
        if not page_id:
            if not quiet:
                print(f"  WARNING: Slide {snum} not found in presentation {pid}")
            continue

        try:
            thumbnail = slides_svc.presentations().pages().getThumbnail(
                presentationId=pid,
                pageObjectId=page_id,
                thumbnailProperties_thumbnailSize="LARGE",
            ).execute()

            fd, local_path = tempfile.mkstemp(suffix=".png", prefix="md2gdoc_")
            os.close(fd)
            temp_files.append(local_path)
            urllib.request.urlretrieve(thumbnail["contentUrl"], local_path)
            exported[key] = local_path
            if not quiet:
                print(f"  Exported Slide {snum}: {local_path}")
        except Exception as e:
            print(f"  WARNING: Failed to export slide {snum} from {pid}: {e}")
            continue

    try:
        # Find and replace placeholders (reverse order to preserve indices)
        doc = get_doc(docs_svc, doc_id)
        body_content = doc["body"]["content"]

        placeholders = []
        for el in body_content:
            if "paragraph" not in el:
                continue
            for elem in el["paragraph"].get("elements", []):
                text = elem.get("textRun", {}).get("content", "")
                m = re.match(r"\{\{DIAGRAM:([^:]+):SLIDE(\d+):([^}]+)\}\}", text)
                if m:
                    placeholders.append({
                        "presentation_id": m.group(1),
                        "slide_num": int(m.group(2)),
                        "label": m.group(3),
                        "start": elem["startIndex"],
                        "end": elem["endIndex"],
                    })

        if not placeholders:
            if not quiet:
                print("  No diagram placeholders found in doc")
            return

        placeholders.sort(key=lambda p: p["start"], reverse=True)

        from googleapiclient.http import MediaFileUpload
        embedded_count = 0
        for ph in placeholders:
            key = (ph["presentation_id"], ph["slide_num"])
            local_path = exported.get(key)
            if not local_path:
                continue

            try:
                file_metadata = {"name": f"diagram_{ph['label'].replace(' ', '_')}.png"}
                media = MediaFileUpload(local_path, mimetype="image/png")
                uploaded = drive_svc.files().create(
                    body=file_metadata, media_body=media, fields="id",
                ).execute()
                file_id = uploaded["id"]

                # Only make world-readable if explicitly requested
                if public_diagrams:
                    drive_svc.permissions().create(
                        fileId=file_id,
                        body={"type": "anyone", "role": "reader"},
                    ).execute()

                img_url = f"https://drive.google.com/uc?id={file_id}"
                reqs = [
                    {"deleteContentRange": {"range": {
                        "startIndex": ph["start"], "endIndex": ph["end"]}}},
                    {"insertInlineImage": {
                        "location": {"index": ph["start"]},
                        "uri": img_url,
                        "objectSize": {"width": {"magnitude": PAGE_WIDTH_PT, "unit": "PT"}},
                    }},
                ]
                batch_send(docs_svc, doc_id, reqs, f"Embed: {ph['label']}", quiet=quiet)
                embedded_count += 1
            except Exception as e:
                print(f"  WARNING: Failed to embed diagram '{ph['label']}': {e}")
                continue

        if not quiet:
            print(f"  Embedded {embedded_count} of {len(placeholders)} diagram(s)")
    finally:
        # Clean up temp files
        for tf in temp_files:
            try:
                os.unlink(tf)
            except OSError:
                pass


# ── Main conversion ───────────────────────────────────────────────────

def convert(md_text: str, doc_id: str = None, title: str = None,
            style: dict | None = None, token_path: str = None,
            embed: bool = True, quiet: bool = False,
            public_diagrams: bool = False) -> str:
    """Convert markdown to a formatted Google Doc.

    Args:
        md_text: Markdown content
        doc_id: Existing doc ID to overwrite (creates new if None)
        title: Document title
        style: Style dict from styles.get_style() (None = no formatting)
        token_path: Path to Google OAuth token JSON
        embed: Whether to embed diagrams from Google Slides
        quiet: Suppress progress output
        public_diagrams: Make uploaded diagram images world-readable (default False)

    Returns:
        Google Doc URL
    """
    from .parser import parse_markdown

    blocks = parse_markdown(md_text)
    if not quiet:
        print(f"Parsed {len(blocks)} blocks")

    creds = get_credentials(token_path)
    docs_svc = get_docs_service(creds)

    # Determine title
    if not title:
        for b in blocks:
            if b["type"] == "heading" and b["level"] == 1:
                plain, _ = strip_md(b["text"])
                title = plain
                break
        if not title:
            title = "Untitled Document"

    if doc_id:
        if not quiet:
            print(f"Clearing doc {doc_id}...")
        end = get_doc_end(docs_svc, doc_id)
        if end > 2:
            docs_svc.documents().batchUpdate(documentId=doc_id, body={"requests": [
                {"deleteContentRange": {"range": {"startIndex": 1, "endIndex": end - 1}}}
            ]}).execute()
    else:
        if not quiet:
            print(f"Creating doc: {title}")
        doc = docs_svc.documents().create(body={"title": title}).execute()
        doc_id = doc["documentId"]
        if not quiet:
            print(f"  Created: {doc_id}")

    # Phase 1: Insert content
    if not quiet:
        print("\nPhase 1: Inserting content...")
    insert_content(docs_svc, doc_id, blocks, style=style, quiet=quiet)

    # Phase 2: Embed diagrams
    has_diagrams = any(b["type"] == "diagram" for b in blocks)
    if has_diagrams and embed:
        if not quiet:
            print("\nPhase 2: Embedding diagrams...")
        embed_diagrams(docs_svc, doc_id, blocks, creds=creds, quiet=quiet,
                       public_diagrams=public_diagrams)

    # Phase 3: Styling
    if style:
        if not quiet:
            print(f"\nPhase 3: Applying style '{style.get('name', 'custom')}'...")
        apply_styling(docs_svc, doc_id, style, quiet=quiet)

        # Phase 4: Readability spacing
        if not quiet:
            print("\nPhase 4: Readability spacing...")
        apply_spacing(docs_svc, doc_id, style, quiet=quiet)

    url = f"https://docs.google.com/document/d/{doc_id}/edit"
    if not quiet:
        print(f"\n=== DONE ===")
        print(f"Doc: {url}")
    return url
