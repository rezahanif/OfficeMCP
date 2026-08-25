# Office Connection Setup Guide

## 1. Prerequisites
- Microsoft Office installed (Word, Excel, PowerPoint, Outlook, MS Project, or Access).
- Windows with COM automation enabled.
- AiConnect Desktop running.

## 2. No Plugin Installation Required

The Office connector talks directly to Office applications via COM — no plugin or add-in to install. Just ensure Office is installed and licensed.

## 3. Connect in AiConnect Desktop

1. Open **AiConnect Desktop** → **MCP Collection**.
2. Find **Office Connector** and click **Enable**.

## 4. Verify Connection

The connector automatically detects installed Office apps. Check with:
```
List all available Office applications on this system.
```

## 5. Start Working

### COM Tools (Windows-only, require running Office app)

| Tool | What it does |
|------|-------------|
| `Launch` | Start or attach to an Office application |
| `Visible` | Show/hide an application window |
| `Quit` | Close an application |
| `ScreenShot` | Capture the application window |
| `Speak` | Text-to-speech via SAPI |
| `AvailableApps` | List all installed Office apps |

### Document CRUD Tools (cross-platform, no Office needed)

**Word (python-docx):**
- `doc_create` — create a new .docx
- `doc_read` — read document content
- `doc_add_paragraph` — add text paragraphs
- `doc_add_heading` — add headings
- `doc_add_table` — insert tables
- `doc_replace_text` — find and replace text
- `doc_get_properties` — read document metadata

**Excel (openpyxl):**
- `xlsx_create` — create a new .xlsx
- `xlsx_read_cells` — read cell ranges
- `xlsx_write_cells` — write cell data
- `xlsx_list_sheets` — list worksheets
- `xlsx_add_sheet` — add a new worksheet
- `xlsx_append_rows` — append data rows
- `xlsx_get_properties` — read workbook metadata

**PowerPoint (python-pptx):**
- `pptx_create` — create a new .pptx
- `pptx_read` — read presentation content
- `pptx_add_slide` — add slides with title/content
- `pptx_get_info` — read presentation metadata

### Example Agent Prompts

```
Create a new Excel workbook at D:/data/report.xlsx with a sales summary: 
headers in row 1 (Product, Q1, Q2, Q3, Q4, Total), then 5 rows of data.
```

```
Open the Word document at D:/contracts/template.docx, add a new paragraph 
with today's date, and save it.
```

```
Create a PowerPoint presentation with 3 slides: title slide, content 
slide with bullet points, and a closing slide.
```

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `OFFICE_MCP_ROOT` | `~/@OfficeMCP` | Working directory for file operations |

## Security Notes

- **RunPython is REMOVED** — no arbitrary code execution.
- Document CRUD uses pure Python libraries (python-docx, openpyxl, python-pptx) — no COM needed.
- COM tools require the target Office app to be installed and running.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `COM connection failed` | Office app not running. Launch it first with `Launch`. |
| `App not installed` | Check `AvailableApps` to see what's installed. |
| `Permission denied` | Run as Administrator or check file permissions. |
