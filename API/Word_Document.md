# Document.Content

## Syntax
expression.Document.Content

## Description
Returns a Range object representing the main document story. Used to read or
modify the entire body text.

# Document.Paragraphs

## Syntax
expression.Document.Paragraphs

## Description
Collection of Paragraph objects. Iterate to read all paragraph text:
for para in doc.Paragraphs: print(para.Range.Text)

# Document.Tables

## Syntax
expression.Document.Tables

## Description
Collection of Table objects. Each Table has Rows and Cells;
Cell(row, column).Range.Text reads a cell's content.

# Document.SaveAs2

## Syntax
expression.Document.SaveAs2(FileName, FileFormat)

## Description
Saves the document. FileFormat 16 = wdFormatDocumentDefault (.docx).

# Document.Fields.Update

## Syntax
expression.Document.Fields.Update

## Description
Updates all fields (TOC, dates, references) in the document.
