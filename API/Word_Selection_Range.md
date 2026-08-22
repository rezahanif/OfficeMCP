# Selection.TypeText

## Syntax
expression.Selection.TypeText(Text)

## Description
Types text at the current selection/cursor position, replacing any selection.

# Selection.Find.Execute

## Syntax
expression.Selection.Find.Execute(FindText, ReplaceWith, Replace)

## Description
Finds and optionally replaces text. Replace: 1=wdReplaceOne, 2=wdReplaceAll.

## Notes (COM)
The MCP connector's doc_replace_text(path, find, replace) does global
replace cross-platform without Word.

# Range.Text

## Syntax
expression.Range.Text

## Description
Read or set the text of a range. Setting replaces the range content.

# Range.Style

## Syntax
expression.Range.Style = "Heading 1"

## Description
Apply a built-in style by name to a range.
