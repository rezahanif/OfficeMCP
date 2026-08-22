# Worksheet.Cells

## Syntax
expression.Worksheet.Cells(row, column)

## Description
Returns a single cell by 1-based row and column numbers. .Value reads/writes
the cell value.

## Notes (COM)
Cross-platform alternative: xlsx_write_cells(path, sheet, start_cell, data).

# Worksheet.Range

## Syntax
expression.Worksheet.Range("A1:C10")

## Description
Returns a Range from an A1-style address. .Value is a 2D array for
multi-cell ranges.

# Worksheet.Copy / Worksheet.Paste

## Syntax
expression.Range.Copy(Destination)

## Description
Copies a range to a destination range or clipboard.

# Worksheet.UsedRange

## Syntax
expression.Worksheet.UsedRange

## Description
The rectangular region containing all non-empty cells. Useful to read the
whole used area in one call.

# Worksheet.Name

## Syntax
expression.Worksheet.Name = "NewName"

## Description
Renames the worksheet.
