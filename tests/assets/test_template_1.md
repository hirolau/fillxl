# Excel Test File - Sheet Descritions

## Sheet: TestSingleCell

### A1

Named Cell (singel_cell) with value "Single Cell Before" formatted in red and bold.

### E2:F3

Named Range (named_range_multiple_cells) spanning multiple cells with values:

E2: "named range multiple cells 1"
F2: "named range multiple cells 3"
E3: "named range multiple cells 2"
F3: "named range multiple cells 4"

### B8:C9

Named Cell (single_cell_merged_cells) with value "Merdged Cell Value". The cell range B8:C9 is merged.

## Sheet: TestTable

### A1:E5

A defined table with columns:

Col1	Col2	ColNotFilled	ColFormula	TestFormattedPct

The entire table has conditional formatting. If a value is seven, the background is green.

Col2 has additional conditional formatting. If a value is between 3 and 4, the background is red. 

### A40

Single filled cell with value TEST NOT OVERWRITE.

## Sheet: SpilledRange

=TableTest[Col1] in B2. Can be used to test if spilled data works after writing.

## Sheet: JustAnotherSheet

### B2

Named Cell (named_range_on_another_sheet) with value "Named range on another sheet." Used to test named ranges that reference a sheet other than the one being filled.

