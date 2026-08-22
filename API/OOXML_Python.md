# docx.Document (python-docx)

## Syntax
Document(path=None)

## Description
Opens or creates a .docx. Document().save(path) creates a new file.
Cross-platform — no Word installation needed.

```python
from docx import Document
doc = Document()
doc.add_heading("Title", level=0)
doc.add_paragraph("Body text")
doc.save("out.docx")
```

# python-docx: add_table

## Syntax
Document.add_table(rows, cols) / table.style = "Table Grid"

## Description
Appends a table. Fill cells via table.rows[i].cells[j].text.

# openpyxl.load_workbook

## Syntax
load_workbook(path, data_only=False)

## Description
Opens an .xlsx. data_only=True returns cached formula results instead of
formula strings.

```python
from openpyxl import load_workbook
wb = load_workbook("book.xlsx")
ws = wb["Sheet1"]
ws["A1"] = 42
wb.save("book.xlsx")
```

# openpyxl: append rows

## Syntax
Worksheet.append(iterable)

## Description
Adds one row at the bottom of the sheet, starting at column A.

# pptx.Presentation (python-pptx)

## Syntax
Presentation(path=None)

## Description
Opens or creates a .pptx. slide_layouts[1] is the standard title+content
layout.

```python
from pptx import Presentation
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "Agenda"
prs.save("deck.pptx")
```

# EMU units (OOXML)

## Syntax
914400 EMU = 1 inch; 12700 EMU = 1 point

## Description
All OOXML shape dimensions are in English Metric Units. python-pptx accepts
EMU ints directly.
