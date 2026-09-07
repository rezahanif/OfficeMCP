# Microsoft Office & Project MCP Connector Guide

The **Office MCP Connector** (`office-mcp`) provides AI assistants with tools to author and edit Word, Excel, and PowerPoint documents, manage Microsoft Project schedules, and control live Microsoft Office desktop applications.

---

## 1. Architectural Overview & Operational Modes

The connector operates in three complementary modes depending on your requirements and system environment:

```
┌─────────────────────────────────────────────────────────────────────────┐
│                           AI Client / Agent                             │
│             (AiConnect Desktop, Claude Desktop, Cursor, VS Code)         │
└────────────────────────────────────┬────────────────────────────────────┘
                                     │ JSON-RPC 2.0 (stdio)
┌────────────────────────────────────▼────────────────────────────────────┐
│                    Office MCP Server (office-mcp)                       │
├──────────────────────────┬──────────────────────────┬───────────────────┤
│ Mode 1: Document CRUD    │ Mode 2: Live Office COM  │ Mode 3: MS Project│
│ (python-docx, openpyxl,  │ (win32com automation)    │ (win32com)        │
│  python-pptx)            │                          │                   │
│ • Cross-Platform         │ • Windows Only           │ • Windows Only    │
│ • NO Office Install Req. │ • Requires MS Office     │ • Requires Project│
│ • Direct .docx/.xlsx/    │ • Launch, Visible, Quit, │ • Tasks, Gantt,   │
│   .pptx file read/write  │   ScreenShot, Speak      │   Resources, Base │
└──────────────────────────┴──────────────────────────┴───────────────────┘
```

1. **Mode 1: Pure OOXML Document Operations (Cross-Platform)**
   - Operates directly on `.docx`, `.xlsx`, and `.pptx` file archives using `python-docx`, `openpyxl`, and `python-pptx`.
   - **Does NOT require Microsoft Office to be installed**. Works on Windows and headless environments.
   - Ideal for generating reports, parsing spreadsheets, logging data, and creating slide decks.

2. **Mode 2: Live Office Desktop Automation (Windows COM)**
   - Connects directly to running or installed Office desktop applications (`Word`, `Excel`, `PowerPoint`, `Outlook`, `Access`, `MSProject`) via Windows COM automation (`win32com`).
   - Requires Windows and licensed Microsoft Office applications.
   - Allows agents to launch applications, toggle visibility, capture screenshots, and perform SAPI text-to-speech.

3. **Mode 3: Microsoft Project Scheduling (Windows COM)**
   - Automates Microsoft Project (`.mpp`) for end-to-end project management.
   - Supports creating schedules, adding tasks with duration strings (`"3d"`, `"2w"`) and predecessors (`"1,2FS+2d"`), tracking percent complete, assigning work resources, saving baselines (0–10), and switching views.

---

## 2. Setup & Client Configuration

The connector communicates over standard MCP JSON-RPC 2.0 via `stdio`.

### A. AiConnect Desktop
1. Open **AiConnect Desktop** → navigate to **Marketplace / MCP Collection**.
2. Locate **Microsoft Office Documents & Automation** (`office-mcp`).
3. Click **Enable**. AiConnect manages the connector process, environment variables, and token lifecycle automatically.

### B. Claude Desktop
Add the connector to your `claude_desktop_config.json`:

- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS / Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "office": {
      "command": "python",
      "args": [
        "C:/path/to/OfficeMCP/run_server.py"
      ],
      "env": {
        "OFFICE_MCP_ROOT": "C:/Users/YourUsername/@OfficeMCP"
      }
    }
  }
}
```

### C. Cursor
Add the connector to your `.cursor/mcp.json` or global Cursor settings:

```json
{
  "mcpServers": {
    "office": {
      "command": "python",
      "args": ["C:/path/to/OfficeMCP/run_server.py"],
      "env": {
        "OFFICE_MCP_ROOT": "C:/Users/YourUsername/@OfficeMCP"
      }
    }
  }
}
```

### D. Direct stdio Execution
You can test the server directly from PowerShell or terminal:

```bash
python run_server.py
```

---

## 3. Filesystem & Path Resolution

File operations adhere to predictable sandboxing and path rules:

- **Relative Paths**: Automatically resolve against `OFFICE_MCP_ROOT`. If `OFFICE_MCP_ROOT` is not set, the connector defaults to `%USERPROFILE%\@OfficeMCP` (or `~/@OfficeMCP`).
  - Example: `doc_create("reports/q1.docx")` creates `%USERPROFILE%\@OfficeMCP\reports\q1.docx`.
- **Absolute Paths**: Fully supported. Passing an absolute path (e.g. `D:/Finance/budget.xlsx` or `C:\Contracts\master.docx`) bypasses root folder resolution and targets the file directly.
- **Parent Directories**: Automatically created on write if they do not exist.

---

## 4. Comprehensive Tool Reference

### A. Word Document Tools (`.docx`) — Pure Python

| Tool | Parameters | Description |
|---|---|---|
| `doc_create` | `path: str` | Create a new empty Word document (`.docx`). |
| `doc_read` | `path: str` | Extract all paragraph text and table contents (as 2D arrays). |
| `doc_add_paragraph` | `path: str`, `text: str`, `style: str = None` | Append a text paragraph with optional style (`"Normal"`, `"Quote"`, `"List Bullet"`). |
| `doc_add_heading` | `path: str`, `text: str`, `level: int = 1` | Add a heading (level `0` = Document Title, `1` = Heading 1, up to `4`). |
| `doc_add_table` | `path: str`, `rows: list[list[str]]`, `headers: list[str] = None` | Append a grid table with optional header row. |
| `doc_replace_text` | `path: str`, `find: str`, `replace: str` | Find and replace placeholder text across all paragraphs and table cells. |
| `doc_get_properties` | `path: str` | Read document metadata: author, title, created/modified dates, paragraph/table counts. |

---

### B. Excel Spreadsheet Tools (`.xlsx`) — Pure Python

| Tool | Parameters | Description |
|---|---|---|
| `xlsx_create` | `path: str` | Create a new workbook with default `"Sheet1"`. |
| `xlsx_read_cells` | `path: str`, `sheet: str = None`, `cell_range: str = None` | Read cell values. `cell_range` accepts `"A1:D10"`, `"B2"`, or `None` (reads all rows). |
| `xlsx_write_cells` | `path: str`, `sheet: str`, `start_cell: str`, `data: list[list[Any]]` | Write a 2D matrix starting at `start_cell` (e.g. `"A1"`). Creates sheet if missing. |
| `xlsx_list_sheets` | `path: str` | List all sheet names in the workbook. |
| `xlsx_add_sheet` | `path: str`, `name: str` | Add a new worksheet. Fails if name already exists. |
| `xlsx_append_rows` | `path: str`, `sheet: str`, `rows: list[list[Any]]` | Append rows to the end of a sheet without overwriting existing data. |
| `xlsx_get_properties`| `path: str` | Read workbook metadata: title, creator, created date, sheet list. |

---

### C. PowerPoint Presentation Tools (`.pptx`) — Pure Python

| Tool | Parameters | Description |
|---|---|---|
| `pptx_create` | `path: str` | Create a new presentation with an initial blank slide. |
| `pptx_read` | `path: str` | Extract all text per slide, including text boxes and tables. |
| `pptx_add_slide` | `path: str`, `title: str`, `content: str = None` | Add a slide with title and body bullet points. |
| `pptx_get_info` | `path: str` | Read presentation metadata: slide count, author, dimensions. |

---

### D. Microsoft Project Tools (`.mpp`) — Windows COM

| Tool | Parameters | Description |
|---|---|---|
| `msp_create` | `path: str` | Create a new blank Project `.mpp` schedule and save it. |
| `msp_open` | `path: str` | Open an existing `.mpp` schedule file. |
| `msp_save` | `path: str = None` | Save the active project (`Save` or `SaveAs`). |
| `msp_close` | `save: bool = True` | Close active project with optional save. |
| `msp_add_task` | `name: str`, `duration: str = None`, `start: str = None`, `predecessors: str = None`, `notes: str = None` | Add an activity. Durations support `"3d"`, `"1w"`, `"4h"`. Predecessors support `"1"`, `"1FS+2d"`. |
| `msp_get_tasks` | `filter_name: str = None`, `field: str = None`, `value: str = None` | List all activities in schedule with durations, start/finish, progress. |
| `msp_update_task`| `unique_id: int`, `name: str = None`, `duration: str = None`, `percent_complete: int = None`, ... | Update task attributes, reschedule, or mark progress. |
| `msp_delete_task`| `unique_id: int` | Delete a task by UniqueID. |
| `msp_add_resource`| `name: str`, `type: str = "Work"`, `standard_rate: float = None`, `email: str = None` | Add human or material resource to project pool. |
| `msp_get_resources`| *(none)* | List all resources with rates and contact details. |
| `msp_delete_resource`| `unique_id: int` | Delete resource by UniqueID. |
| `msp_assign_resource`| `task_unique_id: int`, `resource_unique_id: int`, `units: float = 1.0` | Assign resource to task (units: `1.0` = 100%). |
| `msp_set_baseline`| `baseline_number: int = 0` | Lock project schedule into baseline (`0` to `10`). |
| `msp_get_project_info`| *(none)* | Summary of project: start/finish dates, task count, resource count. |
| `msp_switch_view`| `view_name: str` | Switch active window view (e.g. `"Gantt Chart"`, `"Task Sheet"`). |
| `msp_list_views` | *(none)* | List all view names available in Project. |

---

### E. Live Office COM Lifecycle & Desktop Control — Windows Only

Valid `app_name` values: `"Word"`, `"Excel"`, `"PowerPoint"`, `"Outlook"`, `"MSProject"`, `"Access"`.

| Tool | Parameters | Description |
|---|---|---|
| `AvailableApps` | *(none)* | Returns list of Office applications installed on the host system. |
| `RunningApps` | *(none)* | Returns list of Office applications currently open and active. |
| `IsAppAvailable` | `app_name: str` | Check if a specific Office app is installed. |
| `Launch` | `app_name: str = "Word"`, `visible: bool = True` | Launch or attach to an Office application window. |
| `Visible` | `app_name: str = "Word"`, `visible: bool = True` | Show or hide an application window. |
| `Quit` | `app_name: str = "Word"`, `force: bool = False` | Terminate or close an Office application. |
| `ScreenShot` | `save_path: str = None` | Capture the primary desktop viewport to PNG. |
| `Speak` | `text: str`, `volume: int = 80`, `rate: int = 0` | SAPI text-to-speech audio output. |
| `Beep` | `frequency: int = 2500`, `duration: int = 1000` | Beep the PC speaker. |
| `RootFolder` | *(none)* | Return current working root folder (`OFFICE_MCP_ROOT`). |
| `IsFileExists` | `sub_file_path: str` | Verify if a file exists relative to the root folder. |
| `DownloadImage`| `url: str`, `save_path: str` | Fetch remote image and save to root folder. |

---

### F. API Discovery & Templates (Layer B)

| Tool | Parameters | Description |
|---|---|---|
| `search_office_api` | `query: str`, `category: str = None` | Search Office API reference docs for methods, classes, and code syntax. |
| `list_office_api_categories` | *(none)* | List available documentation sections and method counts. |
| `list_templates` | *(none)* | List verified document templates (e.g. `word_report_skeleton`, `excel_data_table`). |
| `load_template` | `template_id: str` | Retrieve the recipe and sequence of tool calls for a template. |
| `get_error_hints` | `error_code: str = None` | Return actionable troubleshooting guidance for errors. |

---

## 5. End-to-End Agent Workflows

### Workflow 1: Generate an Executive Word Report (.docx)

An agent can author structured Word documents without Office running:

```json
// 1. Create document
{"name": "doc_create", "arguments": {"path": "Q3_Performance_Report.docx"}}

// 2. Add Title & Subheading
{"name": "doc_add_heading", "arguments": {"path": "Q3_Performance_Report.docx", "text": "Q3 Performance Report", "level": 0}}
{"name": "doc_add_heading", "arguments": {"path": "Q3_Performance_Report.docx", "text": "Executive Summary", "level": 1}}

// 3. Add text paragraphs
{"name": "doc_add_paragraph", "arguments": {"path": "Q3_Performance_Report.docx", "text": "During Q3, overall revenue increased by 14% quarter-over-quarter across all product lines."}}

// 4. Add structured table with data
{"name": "doc_add_table", "arguments": {
  "path": "Q3_Performance_Report.docx",
  "headers": ["Business Unit", "Target (k$)", "Actual (k$)", "Growth"],
  "rows": [
    ["Enterprise Cloud", "1,200", "1,380", "+15%"],
    ["Edge Solutions", "850", "910", "+7%"],
    ["Consulting", "400", "440", "+10%"]
  ]
}}

// 5. Populate template placeholders if reusing boilerplate
{"name": "doc_replace_text", "arguments": {
  "path": "Q3_Performance_Report.docx",
  "find": "{{REVIEW_DATE}}",
  "replace": "September 2026"
}}
```

---

### Workflow 2: Financial Workbook & Logging in Excel (.xlsx)

```json
// 1. Create workbook
{"name": "xlsx_create", "arguments": {"path": "finance/quarterly_model.xlsx"}}

// 2. Write headers and financial figures into Sheet1
{"name": "xlsx_write_cells", "arguments": {
  "path": "finance/quarterly_model.xlsx",
  "sheet": "Revenue",
  "start_cell": "A1",
  "data": [
    ["Month", "Gross Revenue", "Operating Cost", "Net Margin"],
    ["July", 340000, 210000, 130000],
    ["August", 365000, 215000, 150000],
    ["September", 410000, 225000, 185000]
  ]
}}

// 3. Add a secondary worksheet for audit trail
{"name": "xlsx_add_sheet", "arguments": {"path": "finance/quarterly_model.xlsx", "name": "AuditLog"}}

// 4. Append audit entries
{"name": "xlsx_append_rows", "arguments": {
  "path": "finance/quarterly_model.xlsx",
  "sheet": "AuditLog",
  "rows": [
    ["2026-09-07T12:00:00Z", "AI Agent", "Initial quarterly data inserted"],
    ["2026-09-07T12:05:00Z", "Financial Analyst", "Approved figures"]
  ]
}}

// 5. Read back a specific range for verification
{"name": "xlsx_read_cells", "arguments": {
  "path": "finance/quarterly_model.xlsx",
  "sheet": "Revenue",
  "cell_range": "A1:D4"
}}
```

---

### Workflow 3: Multi-Slide Briefing Deck in PowerPoint (.pptx)

```json
// 1. Create deck
{"name": "pptx_create", "arguments": {"path": "presentations/product_briefing.pptx"}}

// 2. Add agenda slide
{"name": "pptx_add_slide", "arguments": {
  "path": "presentations/product_briefing.pptx",
  "title": "Agenda",
  "content": "• Market Landscape\n• Architecture Overview\n• Milestones & Timeline\n• Q&A"
}}

// 3. Add content slide
{"name": "pptx_add_slide", "arguments": {
  "path": "presentations/product_briefing.pptx",
  "title": "Architecture Overview",
  "content": "AiConnect provides a unified gateway with sub-process isolation, zero-trust token minting, and cross-platform document authoring."
}}

// 4. Verify slide contents
{"name": "pptx_read", "arguments": {"path": "presentations/product_briefing.pptx"}}
```

---

### Workflow 4: Full Schedule Lifecycle in Microsoft Project (.mpp)

Requires Windows and Microsoft Project installed:

```json
// 1. Create new schedule file
{"name": "msp_create", "arguments": {"path": "schedules/ai_migration.mpp"}}

// 2. Add Project tasks with durations and dependencies
{"name": "msp_add_task", "arguments": {
  "name": "Architecture Design",
  "duration": "5d"
}}
{"name": "msp_add_task", "arguments": {
  "name": "Backend Implementation",
  "duration": "10d",
  "predecessors": "1FS"
}}
{"name": "msp_add_task", "arguments": {
  "name": "QA & Load Testing",
  "duration": "5d",
  "predecessors": "2FS"
}}

// 3. Add resources
{"name": "msp_add_resource", "arguments": {"name": "Lead Engineer", "type": "Work", "standard_rate": 120.0}}
{"name": "msp_add_resource", "arguments": {"name": "QA Specialist", "type": "Work", "standard_rate": 95.0}}

// 4. Assign resources to tasks
{"name": "msp_assign_resource", "arguments": {"task_unique_id": 2, "resource_unique_id": 1, "units": 1.0}}
{"name": "msp_assign_resource", "arguments": {"task_unique_id": 3, "resource_unique_id": 2, "units": 1.0}}

// 5. Lock baseline schedule
{"name": "msp_set_baseline", "arguments": {"baseline_number": 0}}

// 6. Save and switch to Gantt Chart
{"name": "msp_save", "arguments": {}}
{"name": "msp_switch_view", "arguments": {"view_name": "Gantt Chart"}}
```

---

### Workflow 5: Live COM Inspection & UI Capture

```json
// 1. Inspect what apps are installed
{"name": "AvailableApps", "arguments": {}}

// 2. Launch Excel and make visible
{"name": "Launch", "arguments": {"app_name": "Excel", "visible": true}}

// 3. Check open apps
{"name": "RunningApps", "arguments": {}}

// 4. Capture screenshot of desktop
{"name": "ScreenShot", "arguments": {"save_path": "audit/excel_live.png"}}

// 5. Notify user via text-to-speech
{"name": "Speak", "arguments": {"text": "Excel automation completed successfully."}}

// 6. Gracefully close application
{"name": "Quit", "arguments": {"app_name": "Excel", "force": false}}
```

---

## 6. Security Notes & Error Recovery

### Security Architecture
- **Zero Arbitrary Execution**: In upstream forks, `RunPython` previously allowed running arbitrary Python code with direct COM access. In this AiConnect connector, **`RunPython` is permanently removed**.
- **Pure Library Safety**: Document operations (Word, Excel, PowerPoint) are executed through strict file-format parsers with zero COM or operating system shell calls.

### Error Recovery & Hints
If a COM or file error occurs, call `get_error_hints`:

```json
{"name": "get_error_hints", "arguments": {"error_code": "com_connection"}}
```

| Common Error | Cause | Recovery Step |
|---|---|---|
| `COMConnectionError` | Office application not running or COM unregistered. | Call `Launch(app_name)` first or ensure Microsoft Office is licensed on Windows. |
| `AppNotInstalledError` | Requested Office app (e.g. MSProject) is not installed. | Call `AvailableApps` to verify installed software. For Word/Excel/PPT, use the pure Python `doc_*`, `xlsx_*`, `pptx_*` tools instead. |
| `FileNotFoundError` | File does not exist at the resolved path. | Check `IsFileExists` or verify `OFFICE_MCP_ROOT`. Parent directories are auto-created when writing new files. |

