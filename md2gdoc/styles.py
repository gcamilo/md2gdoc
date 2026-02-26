"""Style presets for Google Docs formatting.

Each style is a dict defining fonts, colors, sizes, and spacing for all
document elements. Pass style=None to skip formatting entirely.

Custom styles: pass a JSON file path via --style-file. The JSON should
follow the same structure as the built-in presets.
"""

import json
from pathlib import Path

# ── Color constants ────────────────────────────────────────────────────

_NAVY = {"red": 0.0, "green": 0.133, "blue": 0.302}
_GOLD = {"red": 0.831, "green": 0.659, "blue": 0.263}
_DARK_GRAY = {"red": 0.2, "green": 0.2, "blue": 0.2}
_CHARCOAL = {"red": 0.13, "green": 0.13, "blue": 0.13}
_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
_LIGHT_BG = {"red": 0.95, "green": 0.95, "blue": 0.97}
_WHITE_BG = {"red": 1.0, "green": 1.0, "blue": 1.0}
_LIGHT_GRAY = {"red": 0.96, "green": 0.96, "blue": 0.96}
_MED_GRAY = {"red": 0.7, "green": 0.7, "blue": 0.7}
_DARK_BG = {"red": 0.26, "green": 0.26, "blue": 0.26}
_BLUE = {"red": 0.11, "green": 0.40, "blue": 0.73}
_LIGHT_BLUE = {"red": 0.85, "green": 0.92, "blue": 0.98}

# ── Built-in presets ───────────────────────────────────────────────────

STYLES = {
    "ib": {
        "name": "Investment Banking",
        "body": {
            "font": "Georgia",
            "size": 11,
            "color": _DARK_GRAY,
            "line_spacing": 120,
            "space_above": 2,
            "space_below": 8,
        },
        "headings": {
            "font": "Georgia",
            "color": _NAVY,
            "sizes": {1: 18, 2: 14, 3: 12, 4: 11},
            "bold": {1: True, 2: True, 3: False, 4: False},
            "space_above": {1: 30, 2: 24, 3: 20, 4: 14},
            "space_below": {1: 8, 2: 8, 3: 6, 4: 4},
        },
        "h2_border": _GOLD,
        "tables": {
            "font": "Arial",
            "size": 9,
            "header_bg": _NAVY,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_BG,
            "row_bg": _WHITE_BG,
            "body_color": _DARK_GRAY,
            "cell_padding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
            "header_padding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
        },
        "blockquote": {
            "border_color": _NAVY,
            "border_width": 3,
            "indent": 18,
        },
        "bullets": {
            "space_above": 1,
            "space_below": 3,
            "line_spacing": 115,
            "list_start_space": 6,
            "list_end_space": 8,
        },
        "margins": {"top": 54, "bottom": 54, "left": 72, "right": 72},
    },

    "clean": {
        "name": "Clean",
        "body": {
            "font": "Arial",
            "size": 11,
            "color": _CHARCOAL,
            "line_spacing": 115,
            "space_above": 2,
            "space_below": 6,
        },
        "headings": {
            "font": "Arial",
            "color": _CHARCOAL,
            "sizes": {1: 20, 2: 16, 3: 13, 4: 11},
            "bold": {1: True, 2: True, 3: True, 4: True},
            "space_above": {1: 28, 2: 22, 3: 18, 4: 12},
            "space_below": {1: 8, 2: 6, 3: 4, 4: 4},
        },
        "h2_border": None,
        "tables": {
            "font": "Arial",
            "size": 10,
            "header_bg": _DARK_BG,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_GRAY,
            "row_bg": _WHITE_BG,
            "body_color": _CHARCOAL,
            "cell_padding": {"top": 3, "bottom": 3, "left": 5, "right": 5},
            "header_padding": {"top": 4, "bottom": 4, "left": 5, "right": 5},
        },
        "blockquote": {
            "border_color": _MED_GRAY,
            "border_width": 2,
            "indent": 18,
        },
        "bullets": {
            "space_above": 1,
            "space_below": 2,
            "line_spacing": 115,
            "list_start_space": 4,
            "list_end_space": 6,
        },
        "margins": {"top": 72, "bottom": 72, "left": 72, "right": 72},
    },

    "corporate": {
        "name": "Corporate",
        "body": {
            "font": "Calibri",
            "size": 11,
            "color": _CHARCOAL,
            "line_spacing": 120,
            "space_above": 2,
            "space_below": 6,
        },
        "headings": {
            "font": "Calibri",
            "color": _BLUE,
            "sizes": {1: 22, 2: 16, 3: 13, 4: 11},
            "bold": {1: True, 2: True, 3: True, 4: False},
            "space_above": {1: 30, 2: 24, 3: 18, 4: 12},
            "space_below": {1: 8, 2: 6, 3: 4, 4: 4},
        },
        "h2_border": _BLUE,
        "tables": {
            "font": "Calibri",
            "size": 10,
            "header_bg": _BLUE,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_BLUE,
            "row_bg": _WHITE_BG,
            "body_color": _CHARCOAL,
            "cell_padding": {"top": 3, "bottom": 3, "left": 5, "right": 5},
            "header_padding": {"top": 4, "bottom": 4, "left": 5, "right": 5},
        },
        "blockquote": {
            "border_color": _BLUE,
            "border_width": 2,
            "indent": 18,
        },
        "bullets": {
            "space_above": 1,
            "space_below": 2,
            "line_spacing": 115,
            "list_start_space": 4,
            "list_end_space": 6,
        },
        "margins": {"top": 72, "bottom": 72, "left": 72, "right": 72},
    },
}


def get_style(name: str | None = None, style_file: str | None = None) -> dict | None:
    """Get a style by name or from a JSON file. Returns None for no styling."""
    if style_file:
        with open(style_file) as f:
            return json.load(f)
    if name is None or name == "none":
        return None
    if name not in STYLES:
        available = ", ".join(sorted(STYLES.keys()))
        raise ValueError(f"Unknown style {name!r}. Available: {available}, none")
    return STYLES[name]


def list_styles() -> list[tuple[str, str]]:
    """Return list of (name, description) for all built-in styles."""
    return [(k, v["name"]) for k, v in STYLES.items()]
