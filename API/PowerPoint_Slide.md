# Slide.Shapes

## Syntax
expression.Slide.Shapes

## Description
Collection of shapes on the slide: text boxes, pictures, tables, charts.
Shape.HasTextFrame / .TextFrame.TextRange.Text reads shape text.

# Shapes.AddTextbox

## Syntax
expression.Slide.Shapes.AddTextbox(Orientation, Left, Top, Width, Height)

## Description
Adds a text box. Coordinates in points (72 per inch).

# Shapes.AddTable

## Syntax
expression.Slide.Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height)

## Description
Adds a table to the slide. Access cells via Table.Cell(row, col).Shape.

# Slide.FollowMasterBackground

## Syntax
expression.Slide.FollowMasterBackground = False

## Description
Allows the slide to have its own background, independent of the master.
