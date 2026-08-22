# Range.Formula

## Syntax
expression.Range.Formula = "=SUM(A1:A10)"

## Description
Sets a formula. Reading returns the formula string; Range.Value returns the
computed result.

# Range.NumberFormat

## Syntax
expression.Range.NumberFormat = "0.00"

## Description
Number format string: "0.00" two decimals, "0%" percent, "yyyy-mm-dd" date,
"#,##0" thousands separator.

# Range.Font

## Syntax
expression.Range.Font.Bold = True / .Size = 12 / .Color = RGB(255,0,0)

## Description
Character formatting for the range.

# Range.AutoFill

## Syntax
expression.Range.AutoFill(Destination, Type)

## Description
Fills Destination based on Source pattern. Type: 0=default, 1=copy.

# Range.Sort

## Syntax
expression.Range.Sort(Key1, Order1, Header)

## Description
Sorts the range. Order1: 1=ascending 2=descending. Header: 1=yes 2=no.
