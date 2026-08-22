# Workbook.Sheets

## Syntax
expression.Workbook.Sheets

## Description
Collection of all sheets (worksheets and chart sheets). Add via Sheets.Add.

# Worksheets.Add

## Syntax
expression.Workbook.Worksheets.Add(Before, After, Count)

## Description
Adds one or more worksheets. Returns the new sheet.
Cross-platform alternative: xlsx_add_sheet(path, name).

# Workbook.Save

## Syntax
expression.Workbook.Save

## Description
Saves the workbook to its current file.

# Workbook.SaveAs

## Syntax
expression.Workbook.SaveAs(FileName, FileFormat)

## Description
Saves to a new file. FileFormat 51 = xlOpenXMLWorkbook (.xlsx).
