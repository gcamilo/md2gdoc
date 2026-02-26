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
_TEAL = {"red": 0.0, "green": 0.51, "blue": 0.51}
_LIGHT_TEAL = {"red": 0.88, "green": 0.96, "blue": 0.96}
_FOREST = {"red": 0.13, "green": 0.37, "blue": 0.22}
_LIGHT_GREEN = {"red": 0.91, "green": 0.96, "blue": 0.91}
_BURGUNDY = {"red": 0.50, "green": 0.0, "blue": 0.13}
_LIGHT_ROSE = {"red": 0.98, "green": 0.93, "blue": 0.94}
_CREAM = {"red": 0.99, "green": 0.97, "blue": 0.93}
_WARM_GRAY = {"red": 0.40, "green": 0.36, "blue": 0.33}
_SLATE = {"red": 0.28, "green": 0.33, "blue": 0.41}
_LIGHT_SLATE = {"red": 0.91, "green": 0.93, "blue": 0.96}
_PURPLE = {"red": 0.40, "green": 0.20, "blue": 0.58}
_LIGHT_PURPLE = {"red": 0.94, "green": 0.91, "blue": 0.98}
_ORANGE = {"red": 0.85, "green": 0.40, "blue": 0.0}
_LIGHT_ORANGE = {"red": 1.0, "green": 0.95, "blue": 0.88}
_BLACK = {"red": 0.0, "green": 0.0, "blue": 0.0}
_NEAR_BLACK = {"red": 0.07, "green": 0.07, "blue": 0.07}
_OFF_WHITE = {"red": 0.98, "green": 0.98, "blue": 0.98}

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

    "academic": {
        "name": "Academic",
        "body": {
            "font": "Times New Roman",
            "size": 12,
            "color": _BLACK,
            "line_spacing": 200,  # Double-spaced
            "space_above": 0,
            "space_below": 0,
        },
        "headings": {
            "font": "Times New Roman",
            "color": _BLACK,
            "sizes": {1: 16, 2: 14, 3: 12, 4: 12},
            "bold": {1: True, 2: True, 3: True, 4: False},
            "space_above": {1: 24, 2: 18, 3: 12, 4: 12},
            "space_below": {1: 6, 2: 6, 3: 6, 4: 6},
        },
        "h2_border": None,
        "tables": {
            "font": "Times New Roman",
            "size": 10,
            "header_bg": _WHITE_BG,
            "header_color": _BLACK,
            "header_bold": True,
            "alt_row_bg": _WHITE_BG,
            "row_bg": _WHITE_BG,
            "body_color": _BLACK,
            "cell_padding": {"top": 2, "bottom": 2, "left": 4, "right": 4},
            "header_padding": {"top": 3, "bottom": 3, "left": 4, "right": 4},
        },
        "blockquote": {
            "border_color": _MED_GRAY,
            "border_width": 1,
            "indent": 36,
        },
        "bullets": {
            "space_above": 0,
            "space_below": 0,
            "line_spacing": 200,
            "list_start_space": 6,
            "list_end_space": 6,
        },
        "margins": {"top": 72, "bottom": 72, "left": 72, "right": 72},
    },

    "memo": {
        "name": "Executive Memo",
        "body": {
            "font": "Garamond",
            "size": 12,
            "color": _NEAR_BLACK,
            "line_spacing": 130,
            "space_above": 2,
            "space_below": 6,
        },
        "headings": {
            "font": "Garamond",
            "color": _BURGUNDY,
            "sizes": {1: 20, 2: 15, 3: 13, 4: 12},
            "bold": {1: True, 2: True, 3: True, 4: False},
            "space_above": {1: 28, 2: 22, 3: 16, 4: 12},
            "space_below": {1: 8, 2: 6, 3: 4, 4: 4},
        },
        "h2_border": _BURGUNDY,
        "tables": {
            "font": "Arial",
            "size": 9,
            "header_bg": _BURGUNDY,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_ROSE,
            "row_bg": _WHITE_BG,
            "body_color": _NEAR_BLACK,
            "cell_padding": {"top": 3, "bottom": 3, "left": 6, "right": 6},
            "header_padding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
        },
        "blockquote": {
            "border_color": _BURGUNDY,
            "border_width": 2,
            "indent": 24,
        },
        "bullets": {
            "space_above": 1,
            "space_below": 2,
            "line_spacing": 125,
            "list_start_space": 6,
            "list_end_space": 8,
        },
        "margins": {"top": 72, "bottom": 72, "left": 90, "right": 72},
    },

    "startup": {
        "name": "Startup / Pitch Deck",
        "body": {
            "font": "Inter",
            "size": 11,
            "color": _SLATE,
            "line_spacing": 130,
            "space_above": 2,
            "space_below": 8,
        },
        "headings": {
            "font": "Inter",
            "color": _NEAR_BLACK,
            "sizes": {1: 24, 2: 18, 3: 14, 4: 11},
            "bold": {1: True, 2: True, 3: True, 4: True},
            "space_above": {1: 36, 2: 28, 3: 20, 4: 14},
            "space_below": {1: 10, 2: 8, 3: 6, 4: 4},
        },
        "h2_border": None,
        "tables": {
            "font": "Inter",
            "size": 10,
            "header_bg": _NEAR_BLACK,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _OFF_WHITE,
            "row_bg": _WHITE_BG,
            "body_color": _SLATE,
            "cell_padding": {"top": 4, "bottom": 4, "left": 8, "right": 8},
            "header_padding": {"top": 5, "bottom": 5, "left": 8, "right": 8},
        },
        "blockquote": {
            "border_color": _NEAR_BLACK,
            "border_width": 3,
            "indent": 18,
        },
        "bullets": {
            "space_above": 2,
            "space_below": 4,
            "line_spacing": 130,
            "list_start_space": 8,
            "list_end_space": 10,
        },
        "margins": {"top": 54, "bottom": 54, "left": 72, "right": 72},
    },

    "consulting": {
        "name": "Management Consulting",
        "body": {
            "font": "Arial",
            "size": 10,
            "color": _CHARCOAL,
            "line_spacing": 115,
            "space_above": 1,
            "space_below": 4,
        },
        "headings": {
            "font": "Arial",
            "color": _TEAL,
            "sizes": {1: 18, 2: 14, 3: 11, 4: 10},
            "bold": {1: True, 2: True, 3: True, 4: True},
            "space_above": {1: 28, 2: 20, 3: 14, 4: 10},
            "space_below": {1: 6, 2: 4, 3: 3, 4: 2},
        },
        "h2_border": _TEAL,
        "tables": {
            "font": "Arial",
            "size": 9,
            "header_bg": _TEAL,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_TEAL,
            "row_bg": _WHITE_BG,
            "body_color": _CHARCOAL,
            "cell_padding": {"top": 2, "bottom": 2, "left": 4, "right": 4},
            "header_padding": {"top": 3, "bottom": 3, "left": 4, "right": 4},
        },
        "blockquote": {
            "border_color": _TEAL,
            "border_width": 2,
            "indent": 14,
        },
        "bullets": {
            "space_above": 0,
            "space_below": 2,
            "line_spacing": 110,
            "list_start_space": 4,
            "list_end_space": 4,
        },
        "margins": {"top": 54, "bottom": 54, "left": 54, "right": 54},
    },

    "legal": {
        "name": "Legal Document",
        "body": {
            "font": "Times New Roman",
            "size": 12,
            "color": _BLACK,
            "line_spacing": 150,
            "space_above": 0,
            "space_below": 6,
        },
        "headings": {
            "font": "Times New Roman",
            "color": _BLACK,
            "sizes": {1: 14, 2: 13, 3: 12, 4: 12},
            "bold": {1: True, 2: True, 3: True, 4: False},
            "space_above": {1: 24, 2: 18, 3: 12, 4: 12},
            "space_below": {1: 6, 2: 6, 3: 6, 4: 6},
        },
        "h2_border": None,
        "tables": {
            "font": "Times New Roman",
            "size": 10,
            "header_bg": _CHARCOAL,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_GRAY,
            "row_bg": _WHITE_BG,
            "body_color": _BLACK,
            "cell_padding": {"top": 2, "bottom": 2, "left": 4, "right": 4},
            "header_padding": {"top": 3, "bottom": 3, "left": 4, "right": 4},
        },
        "blockquote": {
            "border_color": _CHARCOAL,
            "border_width": 1,
            "indent": 36,
        },
        "bullets": {
            "space_above": 0,
            "space_below": 3,
            "line_spacing": 150,
            "list_start_space": 6,
            "list_end_space": 6,
        },
        "margins": {"top": 72, "bottom": 72, "left": 90, "right": 72},
    },

    "creative": {
        "name": "Creative Brief",
        "body": {
            "font": "Lato",
            "size": 11,
            "color": _WARM_GRAY,
            "line_spacing": 140,
            "space_above": 2,
            "space_below": 8,
        },
        "headings": {
            "font": "Playfair Display",
            "color": _PURPLE,
            "sizes": {1: 26, 2: 18, 3: 14, 4: 12},
            "bold": {1: False, 2: False, 3: True, 4: True},
            "space_above": {1: 36, 2: 28, 3: 20, 4: 14},
            "space_below": {1: 10, 2: 8, 3: 6, 4: 4},
        },
        "h2_border": _PURPLE,
        "tables": {
            "font": "Lato",
            "size": 10,
            "header_bg": _PURPLE,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_PURPLE,
            "row_bg": _WHITE_BG,
            "body_color": _WARM_GRAY,
            "cell_padding": {"top": 4, "bottom": 4, "left": 8, "right": 8},
            "header_padding": {"top": 5, "bottom": 5, "left": 8, "right": 8},
        },
        "blockquote": {
            "border_color": _PURPLE,
            "border_width": 3,
            "indent": 24,
        },
        "bullets": {
            "space_above": 2,
            "space_below": 4,
            "line_spacing": 135,
            "list_start_space": 8,
            "list_end_space": 8,
        },
        "margins": {"top": 72, "bottom": 72, "left": 72, "right": 72},
    },

    "report": {
        "name": "Annual Report",
        "body": {
            "font": "Source Serif Pro",
            "size": 11,
            "color": _CHARCOAL,
            "line_spacing": 125,
            "space_above": 2,
            "space_below": 6,
        },
        "headings": {
            "font": "Source Sans Pro",
            "color": _FOREST,
            "sizes": {1: 22, 2: 16, 3: 13, 4: 11},
            "bold": {1: True, 2: True, 3: True, 4: False},
            "space_above": {1: 32, 2: 24, 3: 18, 4: 12},
            "space_below": {1: 8, 2: 6, 3: 4, 4: 4},
        },
        "h2_border": _FOREST,
        "tables": {
            "font": "Source Sans Pro",
            "size": 9,
            "header_bg": _FOREST,
            "header_color": _WHITE,
            "header_bold": True,
            "alt_row_bg": _LIGHT_GREEN,
            "row_bg": _WHITE_BG,
            "body_color": _CHARCOAL,
            "cell_padding": {"top": 3, "bottom": 3, "left": 6, "right": 6},
            "header_padding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
        },
        "blockquote": {
            "border_color": _FOREST,
            "border_width": 2,
            "indent": 18,
        },
        "bullets": {
            "space_above": 1,
            "space_below": 3,
            "line_spacing": 120,
            "list_start_space": 6,
            "list_end_space": 6,
        },
        "margins": {"top": 72, "bottom": 72, "left": 72, "right": 72},
    },

    "minimal": {
        "name": "Minimal",
        "body": {
            "font": "Roboto",
            "size": 10,
            "color": _DARK_GRAY,
            "line_spacing": 115,
            "space_above": 1,
            "space_below": 4,
        },
        "headings": {
            "font": "Roboto",
            "color": _DARK_GRAY,
            "sizes": {1: 18, 2: 14, 3: 12, 4: 10},
            "bold": {1: False, 2: False, 3: True, 4: True},
            "space_above": {1: 24, 2: 18, 3: 12, 4: 10},
            "space_below": {1: 4, 2: 4, 3: 3, 4: 2},
        },
        "h2_border": None,
        "tables": {
            "font": "Roboto",
            "size": 9,
            "header_bg": _LIGHT_GRAY,
            "header_color": _DARK_GRAY,
            "header_bold": True,
            "alt_row_bg": _WHITE_BG,
            "row_bg": _WHITE_BG,
            "body_color": _DARK_GRAY,
            "cell_padding": {"top": 2, "bottom": 2, "left": 4, "right": 4},
            "header_padding": {"top": 3, "bottom": 3, "left": 4, "right": 4},
        },
        "blockquote": {
            "border_color": _LIGHT_GRAY,
            "border_width": 2,
            "indent": 14,
        },
        "bullets": {
            "space_above": 0,
            "space_below": 2,
            "line_spacing": 110,
            "list_start_space": 3,
            "list_end_space": 4,
        },
        "margins": {"top": 54, "bottom": 54, "left": 54, "right": 54},
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
