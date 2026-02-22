from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas

OUTPUT_PATH = Path("output/pdf/accessible-spreadsheet-component-summary.pdf")
OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

PAGE_WIDTH, PAGE_HEIGHT = letter
MARGIN = 42
CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN)

c = canvas.Canvas(str(OUTPUT_PATH), pagesize=letter)
y = PAGE_HEIGHT - MARGIN


def draw_title(text: str) -> None:
    global y
    c.setFont("Helvetica-Bold", 18)
    c.drawString(MARGIN, y, text)
    y -= 22


def draw_heading(text: str) -> None:
    global y
    c.setFont("Helvetica-Bold", 11.5)
    c.drawString(MARGIN, y, text)
    y -= 14


def draw_paragraph(text: str, font_size: float = 9.8, line_height: float = 12) -> None:
    global y
    c.setFont("Helvetica", font_size)
    lines = simpleSplit(text, "Helvetica", font_size, CONTENT_WIDTH)
    for line in lines:
        c.drawString(MARGIN, y, line)
        y -= line_height


def draw_bullets(items: list[str], font_size: float = 9.6, line_height: float = 11.5) -> None:
    global y
    c.setFont("Helvetica", font_size)
    bullet_indent = 12
    text_width = CONTENT_WIDTH - bullet_indent

    for item in items:
        wrapped = simpleSplit(item, "Helvetica", font_size, text_width)
        if not wrapped:
            continue
        c.drawString(MARGIN, y, "- " + wrapped[0])
        y -= line_height
        for line in wrapped[1:]:
            c.drawString(MARGIN + bullet_indent, y, line)
            y -= line_height


def draw_numbered(items: list[str], font_size: float = 9.6, line_height: float = 11.5) -> None:
    global y
    c.setFont("Helvetica", font_size)

    for idx, item in enumerate(items, start=1):
        prefix = f"{idx}. "
        prefix_width = c.stringWidth(prefix, "Helvetica", font_size)
        text_width = CONTENT_WIDTH - prefix_width
        wrapped = simpleSplit(item, "Helvetica", font_size, text_width)
        if not wrapped:
            continue

        c.drawString(MARGIN, y, prefix + wrapped[0])
        y -= line_height
        for line in wrapped[1:]:
            c.drawString(MARGIN + prefix_width, y, line)
            y -= line_height


def section_gap() -> None:
    global y
    y -= 6


draw_title("Accessible Spreadsheet Component - One Page Summary")

draw_heading("What it is")
draw_paragraph(
    "<y11n-spreadsheet> is a Lit 3.0 + TypeScript web component that provides an accessible "
    "spreadsheet-style grid in the browser. It implements WAI-ARIA grid semantics with keyboard-first "
    "interaction for data entry and calculation workflows."
)
section_gap()

draw_heading("What use cases it covers")
draw_paragraph(
    "Primary use case: embedding a spreadsheet-like UI in web apps where users need to navigate, edit, "
    "select ranges, and run formulas while preserving keyboard and screen-reader accessibility."
)
section_gap()

draw_heading("What it does")
draw_bullets(
    [
        "Implements WAI-ARIA grid roles, roving tabindex, and aria-live announcements for active cell feedback.",
        "Supports Excel-like keyboard controls: arrows, Tab/Shift+Tab, Enter, Escape, Delete/Backspace, and Ctrl/Cmd shortcuts.",
        "Provides in-cell editing via Enter, double-click, or type-to-edit, with optional read-only mode.",
        "Evaluates formulas with cell refs, ranges, arithmetic, comparisons, string concatenation, and built-in functions.",
        "Allows custom formula extension through registerFunction(name, fn).",
        "Handles copy/cut/paste via TSV clipboard serialization with bounds-aware paste behavior.",
        "Uses virtual row rendering plus scroll spacers to keep large datasets responsive.",
    ]
)
section_gap()

draw_heading("How it works")
draw_bullets(
    [
        "Grid data is a sparse Map keyed by \"row:col\" (CellData stores rawValue, displayValue, and type).",
        "SelectionManager tracks anchor/head cells and normalized ranges for keyboard and pointer selection.",
        "FormulaEngine parses and evaluates formulas, then recalculates formula cells after edits or data changes.",
        "ClipboardManager converts selected ranges to TSV and parses pasted TSV into bounded cell updates.",
        "The component dispatches cell-change, selection-change, and data-change events for integration hooks.",
    ]
)
section_gap()

draw_heading("How to run/use (minimal)")
draw_numbered(
    [
        "Install dependencies: npm install",
        "Run the local demo/dev app: npm run dev",
        "Register the component in your app (repo pattern): import './src/index.ts'",
        "Place the element in markup: <y11n-spreadsheet rows=\"50\" cols=\"26\"></y11n-spreadsheet>",
        "Provide data through setData(...) or the data property; listen for cell-change/selection-change/data-change events.",
        "Required Node.js version: Not found in repo.",
        "Published package install and production import instructions: Not found in repo.",
    ]
)

if y < MARGIN:
    raise RuntimeError(f"Content overflowed one page (y={y}).")

c.save()
print(str(OUTPUT_PATH))
