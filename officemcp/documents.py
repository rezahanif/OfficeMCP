"""
Office OOXML document operations — cross-platform document manipulation.

Three surfaces, three libraries, zero COM:

  Word (.docx)       → python-docx    → Document(path)
  Excel (.xlsx/.xlsm)→ openpyxl       → Workbook / load_workbook
  PowerPoint (.pptx) → python-pptx   → Presentation(path)

All file paths are relative to OfficeMCP's root folder (Officer.FilePath),
same convention as the rest of the connector. Absolute paths also accepted.

Security posture: pure file-format library operations. No exec, no COM, no
network. The caller is responsible for path sandboxing (OfficeMCP root
folder is the default fence).
"""

import os
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from openpyxl import Workbook, load_workbook
from pptx import Presentation


# ------------------------------------------------------------------ utils
def _resolve(path: str) -> str:
    """Resolve a relative path against OfficeMCP's root folder."""
    if os.path.isabs(path):
        return path
    # Import lazily so the module is importable without Officer ever loading
    from officemcp.Officer import TheOfficer

    tmp = TheOfficer()
    default_root = os.path.join(os.path.expanduser("~"), "@OfficeMCP")
    root = getattr(tmp, "_default_folder", None) or os.environ.get("OFFICE_MCP_ROOT") or default_root
    return os.path.join(root, path)


def _ensure_parent(path: str) -> None:
    parent = os.path.dirname(path)
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)


# ------------------------------------------------------------------ Word
def doc_create(path: str) -> dict:
    """Create a new .docx file."""
    full = _resolve(path)
    _ensure_parent(full)
    Document().save(full)
    return {"success": True, "path": full}


def doc_read(path: str) -> dict:
    """Read all paragraph texts from a .docx file."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    doc = Document(full)
    paras = [p.text for p in doc.paragraphs]
    tables: list[list[list[str]]] = []
    for tbl in doc.tables:
        rows = []
        for row in tbl.rows:
            rows.append([cell.text for cell in row.cells])
        tables.append(rows)
    return {
        "success": True,
        "path": full,
        "paragraphs": paras,
        "table_count": len(tables),
        "tables": tables,
        "paragraph_count": len(paras),
    }


def doc_add_paragraph(path: str, text: str, style: str | None = None) -> dict:
    """Append a paragraph to a .docx file."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    doc = Document(full)
    para = doc.add_paragraph(text)
    if style:
        para.style = style
    doc.save(full)
    return {"success": True, "path": full, "text": text}


def doc_add_heading(path: str, text: str, level: int = 1) -> dict:
    """Add a heading (level 0–4 maps to Heading 1–Heading 4) to a .docx file."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    doc = Document(full)
    doc.add_heading(text, level=level)
    doc.save(full)
    return {"success": True, "path": full, "text": text, "level": level}


def doc_add_table(path: str, rows: list[list[str]], headers: list[str] | None = None) -> dict:
    """Append a table to a .docx file."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    doc = Document(full)
    data_rows = rows
    if headers is not None:
        data_rows = [headers, *rows]
    if not data_rows:
        raise ValueError("No rows provided")
    ncols = max(len(r) for r in data_rows)
    table = doc.add_table(rows=len(data_rows), cols=ncols)
    table.style = "Table Grid"
    for i, row in enumerate(data_rows):
        cells = table.rows[i].cells
        for j, val in enumerate(row):
            if j < len(cells):
                cells[j].text = str(val)
    doc.save(full)
    return {"success": True, "path": full, "rows": len(data_rows), "cols": ncols}


def doc_replace_text(path: str, find: str, replace: str) -> dict:
    """Replace all occurrences of find with replace across paragraphs and tables."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    doc = Document(full)
    replaced = 0
    for para in doc.paragraphs:
        if find in para.text:
            count = 0
            for run in para.runs:
                if find in run.text:
                    c = run.text.count(find)
                    run.text = run.text.replace(find, replace)
                    count += c
            # paragraph.text view may interleave runs; re-scan via full text if still present
            if find in para.text:
                # collapse all runs: replace via paragraph text -> overwrite first run, clear rest
                full_text = para.text.replace(find, replace)
                # naive run collapse
                para.runs[0].text = full_text
                for r in para.runs[1:]:
                    r.text = ""
                replaced += 1
            else:
                replaced += count
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                if find in cell.text:
                    cell.text = cell.text.replace(find, replace)
                    replaced += 1
                for para in cell.paragraphs:
                    for run in para.runs:
                        if find in run.text:
                            run.text = run.text.replace(find, replace)
                            replaced += 1
                    if find in para.text:
                        para.runs[0].text = para.text.replace(find, replace)
                        for r in para.runs[1:]:
                            r.text = ""
                        replaced += 1
    doc.save(full)
    return {"success": True, "path": full, "replacements": replaced}


def doc_get_properties(path: str) -> dict:
    """Read core properties of a .docx file."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    doc = Document(full)
    cp = doc.core_properties
    return {
        "success": True,
        "path": full,
        "title": cp.title,
        "author": cp.author,
        "subject": cp.subject,
        "keywords": cp.keywords,
        "created": str(cp.created) if cp.created else None,
        "modified": str(cp.modified) if cp.modified else None,
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
    }


# ------------------------------------------------------------------ Excel
def xlsx_create(path: str) -> dict:
    """Create a new .xlsx workbook with one default sheet."""
    full = _resolve(path)
    _ensure_parent(full)
    wb = Workbook()
    wb.active.title = "Sheet1"
    wb.save(full)
    return {"success": True, "path": full, "sheets": wb.sheetnames}


def xlsx_read_cells(path: str, sheet: str | None = None, cell_range: str | None = None) -> dict:
    """Read cell values from a workbook. If cell_range is None, reads all non-empty rows."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    wb = load_workbook(full, data_only=True)
    name = sheet or wb.sheetnames[0]
    if name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {name} (available: {wb.sheetnames})")
    ws = wb[name]
    if cell_range is not None:
        rng = ws[cell_range]
        if isinstance(rng, (list, tuple)):
            # multi-cell
            data = [[c.value for c in row] if hasattr(row, "__iter__") else [row.value] for row in rng]
            flat = any(isinstance(v, list) for v in data)
            if not flat:
                data = [data]
        else:
            data = [[rng.value]]
    else:
        data = []
        for row in ws.iter_rows(values_only=True):
            # keep rows that are not entirely None
            if any(v is not None for v in row):
                data.append(list(row))
    return {"success": True, "path": full, "sheet": name, "data": data, "sheets": wb.sheetnames}


def xlsx_write_cells(path: str, sheet: str, start_cell: str, data: list[list[Any]]) -> dict:
    """Write a 2D array starting at start_cell on the named sheet."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    wb = load_workbook(full)
    if sheet not in wb.sheetnames:
        wb.create_sheet(title=sheet)
    ws = wb[sheet]
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

    col_name, row0 = coordinate_from_string(start_cell)
    col0 = column_index_from_string(col_name)
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            ws.cell(row=row0 + r, column=col0 + c, value=val)
    wb.save(full)
    return {"success": True, "path": full, "sheet": sheet, "start": start_cell, "written": len(data) * max((len(r) for r in data), default=0)}


def xlsx_list_sheets(path: str) -> dict:
    """List all sheet names in a workbook."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    wb = load_workbook(full, data_only=True)
    return {"success": True, "path": full, "sheets": wb.sheetnames}


def xlsx_add_sheet(path: str, name: str) -> dict:
    """Add a new sheet to a workbook."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    wb = load_workbook(full)
    if name in wb.sheetnames:
        raise ValueError(f"Sheet already exists: {name}")
    wb.create_sheet(title=name)
    wb.save(full)
    return {"success": True, "path": full, "sheets": wb.sheetnames}


def xlsx_append_rows(path: str, sheet: str, rows: list[list[Any]]) -> dict:
    """Append rows at the end of a sheet."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    wb = load_workbook(full)
    name = sheet or wb.sheetnames[0]
    if name not in wb.sheetnames:
        wb.create_sheet(title=name)
    ws = wb[name]
    for row in rows:
        ws.append(row)
    wb.save(full)
    return {"success": True, "path": full, "sheet": name, "appended": len(rows)}


def xlsx_get_properties(path: str) -> dict:
    """Read workbook metadata."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    wb = load_workbook(full, data_only=True)
    props = wb.properties
    return {
        "success": True,
        "path": full,
        "title": props.title,
        "creator": props.creator,
        "created": str(props.created) if props.created else None,
        "sheets": wb.sheetnames,
        "sheet_count": len(wb.sheetnames),
    }


# ------------------------------------------------------------------ PowerPoint
def pptx_create(path: str) -> dict:
    """Create a new presentation with one title slide."""
    full = _resolve(path)
    _ensure_parent(full)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    if slide.shapes.title:
        slide.shapes.title.text = ""
    prs.save(full)
    return {"success": True, "path": full, "slides": 1}


def pptx_read(path: str) -> dict:
    """Extract text from all slides."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    prs = Presentation(full)
    slides: list[dict] = []
    for i, sl in enumerate(prs.slides):
        texts: list[str] = []
        for shape in sl.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    texts.append(txt)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        txt = cell.text.strip()
                        if txt:
                            texts.append(txt)
        slides.append({"index": i, "texts": texts})
    return {"success": True, "path": full, "slides": slides, "slide_count": len(slides)}


def pptx_add_slide(path: str, title: str, content: str | None = None) -> dict:
    """Add a new slide with a title and optional body content."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    prs = Presentation(full)
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # title + content
    if slide.shapes.title:
        slide.shapes.title.text = title
    if content:
        for ph in slide.placeholders:
            if ph.placeholder_format.type == 2:  # BODY
                ph.text = content
                break
        else:
            # no body placeholder — add a textbox
            left = int(0.5 * 914400)
            top = int(1.5 * 914400)
            w = int(9 * 914400)
            h = int(5 * 914400)
            box = slide.shapes.add_textbox(left, top, w, h)
            box.text_frame.text = content
            box.text_frame.word_wrap = True
    prs.save(full)
    return {"success": True, "path": full, "slide_count": len(prs.slides), "title": title}


def pptx_get_info(path: str) -> dict:
    """Read presentation metadata and slide count."""
    full = _resolve(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    prs = Presentation(full)
    cp = prs.core_properties
    return {
        "success": True,
        "path": full,
        "title": cp.title,
        "author": cp.author,
        "created": str(cp.created) if cp.created else None,
        "slides": len(prs.slides),
        "slide_width": prs.slide_width,  # EMU (914400 = 1 inch)
        "slide_height": prs.slide_height,
    }
