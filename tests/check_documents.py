"""Document CRUD tools validation — runs on ANY OS (no COM, no Office).

Round-trip tests: create → modify → read → verify for Word, Excel,
PowerPoint. Exercises the real documents.py module end-to-end.
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# Isolate the root folder so tests never touch ~/@OfficeMCP
_tmp_root = tempfile.mkdtemp(prefix="officemcp-docs-")
os.environ["OFFICE_MCP_ROOT"] = _tmp_root

results = []


def check(name, cond):
    results.append((name, bool(cond)))
    print(("PASS" if cond else "FAIL"), name)


from officemcp import documents as d

# ── Word ──────────────────────────────────────────────────────────────
r = d.doc_create("test.docx")
check("doc_create", r["success"] and os.path.exists(r["path"]))

d.doc_add_heading("test.docx", "Report Title", level=0)
d.doc_add_paragraph("test.docx", "First paragraph with ALPHA placeholder.")
d.doc_add_paragraph("test.docx", "Second paragraph.")
d.doc_add_table("test.docx", rows=[["r1c1", "r1c2"]], headers=["H1", "H2"])

read = d.doc_read("test.docx")
check("doc_read: heading present", any("Report Title" in p for p in read["paragraphs"]))
check("doc_read: table extracted", read["table_count"] == 1 and read["tables"][0][0] == ["H1", "H2"])

rep = d.doc_replace_text("test.docx", "ALPHA", "BETA")
read2 = d.doc_read("test.docx")
check("doc_replace: replaced", rep["replacements"] >= 1 and any("BETA" in p for p in read2["paragraphs"]))
check("doc_replace: original gone", not any("ALPHA" in p for p in read2["paragraphs"]))

props = d.doc_get_properties("test.docx")
check("doc_props: counts", props["paragraphs"] >= 3 and props["tables"] == 1)

# ── Excel ─────────────────────────────────────────────────────────────
r = d.xlsx_create("book.xlsx")
check("xlsx_create", r["success"] and r["sheets"] == ["Sheet1"])

d.xlsx_write_cells("book.xlsx", "Sheet1", "A1", [["name", "qty"], ["widget", 5], ["gadget", 12]])
data = d.xlsx_read_cells("book.xlsx")
check("xlsx_read: all rows", data["data"][0] == ["name", "qty"] and data["data"][2] == ["gadget", 12])

rng = d.xlsx_read_cells("book.xlsx", cell_range="A1:A3")
check("xlsx_read: scoped range", rng["data"] == [["name"], ["widget"], ["gadget"]])

d.xlsx_add_sheet("book.xlsx", "Summary")
sheets = d.xlsx_list_sheets("book.xlsx")
check("xlsx_add_sheet", sheets["sheets"] == ["Sheet1", "Summary"])

try:
    d.xlsx_add_sheet("book.xlsx", "Summary")
    check("xlsx_add_sheet dup rejected", False)
except ValueError:
    check("xlsx_add_sheet dup rejected", True)

d.xlsx_append_rows("book.xlsx", "Log", [["2026-01-01", "created"]])
log = d.xlsx_read_cells("book.xlsx", sheet="Log")
check("xlsx_append: creates + writes", log["data"] == [["2026-01-01", "created"]])

wrote = d.xlsx_write_cells("book.xlsx", "NewSheet", "B2", [[1, 2]])
newvals = d.xlsx_read_cells("book.xlsx", sheet="NewSheet", cell_range="B2:C2")
check("xlsx_write: auto-create sheet", wrote["success"] and newvals["data"] == [[1, 2]])

props = d.xlsx_get_properties("book.xlsx")
check("xlsx_props", props["sheet_count"] == 4)  # Sheet1, Summary, Log, NewSheet

# ── PowerPoint ────────────────────────────────────────────────────────
r = d.pptx_create("deck.pptx")
check("pptx_create", r["success"] and r["slides"] == 1)

r2 = d.pptx_add_slide("deck.pptx", "Agenda", "1. Intro\n2. Details\n3. Q&A")
check("pptx_add_slide", r2["slide_count"] == 2)

content = d.pptx_read("deck.pptx")
all_text = " ".join(t for s in content["slides"] for t in s["texts"])
check("pptx_read: title found", "Agenda" in all_text)
check("pptx_read: body found", "Details" in all_text)

info = d.pptx_get_info("deck.pptx")
check("pptx_info: slide count", info["slides"] == 2)
check("pptx_info: EMU width sane", info["slide_width"] > 1000000)

failed = [n for n, ok in results if not ok]
print(f"\n{len(results) - len(failed)}/{len(results)} document CRUD checks passed")
sys.exit(1 if failed else 0)
