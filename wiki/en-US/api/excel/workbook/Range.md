---
title: "Range"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Range.html
updated: 1789700916000
hash: DtQbCQWA
---

# Range

Represents a single cell, a row, a column, or a contiguous block of selected cells.

A Range is a subset of cells on a worksheet that you can read or modify together.

Below is a short example. See the lists below for the Range object's properties and methods.

```js
// Select cells A1:A10
let range = Application.Range('A1:A10')
// Set A1:A10 to row numbers
for (let i = 1; i <= 10; i++) {
  range.Item(i).Value = i.toString()
}
// Print
printText(range) // 1...10
// Prefix and suffix each cell
range.Each(cell => {
  cell.Value = 'prefix_' + cell.Text + '_suffix'
})
printText(range) //prefix_1_suffix...prefix_10_suffix

function printText(r) {
  r.Each(cell => {
    console.log(cell.Text)
  })
}
```

#### Properties

| Property | Data type | Description |
|---|---|---|
| [Count](#count) | Number | Number of cells in the range |
| [Text](#text) | String | [Read-only] Formatted text as shown in the cell |
| [Value/Value2](#value) | any/[][]any | Get or set cell values |
| [FormatConditions](#formatConditions) | [FormatConditions](/docs/en-US/api/excel/workbook/FormatConditions.html) | Conditional formatting for the range |
| [Formula](#formula) | String | A1-style formula (implicit intersection) |
| [FormulaArray](#formulaarray) | String | Get or set an array formula for the range |
| [NumberFormat](#numberformat) | String | Get or set number format |
| [Hidden](#hidden) | Boolean | Hide or show row or column |
| [Interior.Color](#interior-color) | String | Interior color as hexadecimal RGB |
| [HorizontalAlignment](#horizontalalignment) | [Enum.XlHAlign](/docs/en-US/api/excel/workbook/Enum.html#xlhalign) | Horizontal alignment |
| [VerticalAlignment](#verticalalignment) | [Enum.XlVAlign](/docs/en-US/api/excel/workbook/Enum.html#xlvalign) | Vertical alignment |
| [WrapText](#wraptext) | Boolean | Wrap text in the range |
| [IndentLevel](#indentlevel) | Number | Cell indent |
| [MergeArea](#mergearea) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Merged area containing the cell |
| [MergeCells](#mergecells) | Boolean | Whether the range contains merged cells |
| [Cells](#cells) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Cells in the range |
| [Rows](#rows) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Rows in the range |
| [Columns](#columns) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Columns in the range |
| [EntireRow](#entirerow) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Full rows that intersect the range |
| [EntireColumn](#entirecolumn) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Full columns that intersect the range |
| [Row](#row) | Number | First row number in the range |
| [RowEnd](#rowend) | Number | Last row number in the range |
| [Column](#column) | Number | First column index in the range |
| [ColumnEnd](#columnend) | Number | Last column index in the range |
| [Borders](#borders) | [Borders](#borders) | Border collection |

#### Methods

| Method | Return type | Description |
|---|---|---|
| [BorderAround()](#borderaround) | undefined | Add a border; set Color, LineStyle, and Weight on the new border |
| [Each()](#each) | undefined | Iterate each cell in the range |
| [Item()](#item) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Cell at the given index within the range |
| [Offset()](#offset) | undefined | Offset the range by rows and columns |
| [Replace()](#replace) | undefined | Replace text inside cells |
| [Delete()](#delete) | undefined | Delete cells, rows, or columns |
| [Insert()](#insert) | undefined | Insert cells, rows, or columns |
| [InsertImage()](#insertimage) | undefined | Insert an image into a cell |
| [Merge()](#merge) | undefined | Merge cells |
| [UnMerge()](#unmerge) | undefined | Unmerge cells |
| [Address()](#address) | String | Address string in macro reference style |
| [AddComment()](#addcomment) | undefined | Add a comment |
| [ClearComments()](#clearcomments) | undefined | Clear comments in the range |
| [Clear()](#clear) | undefined | Clear values and formats in the range |
| [ClearContents()](#clearcontents) | undefined | Clear values only |
| [ClearFormats()](#clearformats) | undefined | Clear formats only |
| [ClearHyperlinks()](#clearhyperlinks) | undefined | Clear hyperlink formatting |
| [Contain()](#contain) | Boolean | Whether two ranges overlap |
| [Copy()](#copy) | Boolean | Copy the range to the clipboard |
| [Cut()](#cut) | Boolean | Cut the range to the destination |
| [PasteSpecial()](#pastespecial) | undefined | Paste from the clipboard with options |
| [FillLeft()](#fillleft) | undefined | Fill left from the rightmost column |
| [FillRight()](#fillright) | undefined | Fill right from the leftmost column |
| [FillDown()](#filldown) | undefined | Fill down from the top row |
| [FillUp()](#fillup) | undefined | Fill up from the bottom row |
| [AutoFill()](#autofill) | undefined | AutoFill a destination range |
| [AutoFilter()](#autofilter) | undefined | Apply AutoFilter criteria to the range |
| [AutoFit()](#autofit) | undefined | Autofit column widths or row heights |
| [Select()](#select) | undefined | Select the range |
| [TextToColumns()](#texttocolumns) | undefined | Split text in one column into multiple columns |

## Count

Count of columns, rows, or cells in the range; defaults to cells

#### Data type

Number - Count of columns, rows, or cells

#### Example

```js
// For a cell range, Range.Count equals Range.Cells.Count
let range = Application.Range('A1:B2')
console.log(range.Count) // 4
console.log(range.Cells.Count) // 4
console.log(range.Rows.Count) // 2
console.log(range.Columns.Count) // 2
```

## Text

[Read-only] Formatted cell text

#### Data type

String - Formatted cell text

#### Example

```js
// Range object
let range = Application.Range('A1')

// [Read-only] Formatted cell text
console.log(range.Text)

// Text is read-only; use Value to change cell content
range.Value = 'NewVaule'
console.log(range.Text) // NewValue
```

## Value/Value2

Get or set cell values; supports reading a range

#### Data type

Single cell: any

Range: [][]any

#### Example

```js
// Value
let range = Application.Range('A1')

range.Value = 'A1'
// Print A1 value
console.log(range.Value) // A1

range = Range('A1:B2')

// Assign A1:B2
range.Value = 'WebOffice'

// Print A1:B2 values
console.log(range.Value) // [["WebOffice","WebOffice"],["WebOffice","WebOffice"]]

// Bulk assign
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]

// Print A1:B2 values
console.log(range.Value) // [["A1","B1"],["A2","B2"]]
```

## FormatConditions

The FormatConditions collection controls conditional formatting in Excel.

#### Data type

[FormatConditions](/docs/en-US/api/excel/workbook/FormatConditions.html) - FormatConditions collection

#### Example

```js
// Get FormatConditions
const formatConditions = Range('A:A').FormatConditions
// Count of format conditions
const count = formatConditions.Count

console.log(count)
```

## Formula

Formula reads and writes cell formulas with different rules than Value.

- When reading: Formula returns the formula; Text returns the displayed value.
- When writing: Formula applies to a single cell; Value supports ranges

#### Data type

String - Formula (implicit intersection)

#### Example

```js
// Set A1=1, B1=2
Application.Range('A1').Value = '1'
Application.Range('B1').Value = '2'

// Set A2 formula =A1+B1
const range = Application.Range('A2')
range.Formula = '=A1+B1'
// Read A2 displayed value: 3
console.log(range.Text) // 3
// Read A2 formula: =A1+B1
console.log(range.Formula) // =A1+B1
```

## FormulaArray

Like Formula but supports setting an array over a range.

#### Data type

String - Array formula for the range

#### Example

```js
let range = Application.Range('A1:B2')

// Set A1:B2 formula text (example)
range.FormulaArray = '=Sum(A1:C3)'

range.Each(function (item) {
  console.log(item.Text) //=Sum(A1:C3), =Sum(A1:C3), =Sum(A1:C3), =Sum(A1:C3)
})
```

## NumberFormat

Get or set the number format for the range.

When getting, if formats differ across the range, the first cell's format is returned.

Common examples you can copy when setting formats in WebOffice:

- General: `G/General` (locale-specific; some locales use a localized general token)
- Number: `0.00_);[Red](0.00)`
- Currency: `￥#,##0.00_);[Red](￥#,##0.00)`
- Accounting: `_ ￥* #,##0.00_ ;_ ￥* -#,##0.00_ ;_ ￥* "-"??_ ;_ @_`
- Short date: `yyyy/m/d;@`
- Long date: locale-specific long date pattern
- Time: `h:mm:ss;@`
- Percent: `0.00%`
- Fraction: `## ?/?`
- Scientific: `0.00E+00`
- Text: `@`
- Thousands separator style: `_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_`

Copy a line above into `Range.NumberFormat`, for example general format: `Range.NumberFormat = 'G/General'`.

#### Data type

String - Number format

#### Example

```js
// Range object
let range = Application.Range('A1')

// Get number format
console.log('Number format:', range.NumberFormat) // General (example locale string)

// Set format to text
range.NumberFormat = '@'
console.log('Number format:', range.NumberFormat) // Text @
```

## Hidden

Hide or show rows or columns; write-only (cannot read previous hidden state).

#### Data type

Boolean - Row or column hidden

#### Example

```js
// Range object
let range = Application.Range('A1')

// Entire column
let entireColumn = range.EntireColumn

// Hide column
entireColumn.Hidden = true
```

## Interior.Color

Interior fill color

#### Data type

Number - Interior color as RGB

#### Example

```js
// Interior object
let interior = Application.Range('A1').Interior

// Get interior color
console.log(interior.Color) // 16777215

// Set interior color to red
interior.Color = RGB(255, 0, 0)
```

## HorizontalAlignment

Horizontal alignment; use values from [Enum.XlHAlign](/docs/en-US/api/excel/workbook/Enum.html#xlhalign).

#### Data type

[Enum.XlHAlign](/docs/en-US/api/excel/workbook/Enum.html#xlhalign) - Horizontal alignment

#### Example

```js
// Range object
let range = Application.Range('A1')

// Print alignment
console.log(range.HorizontalAlignment) // 1

// Set alignment: center
range.HorizontalAlignment = Application.Enum.XlHAlign.xlHAlignCenter

// Print alignment
console.log(range.HorizontalAlignment) // -4108 (center)
```

## VerticalAlignment

Vertical alignment; use values from [Enum.XlVAlign](/docs/en-US/api/excel/workbook/Enum.html#xlvalign).

#### Data type

[Enum.XlVAlign](/docs/en-US/api/excel/workbook/Enum.html#xlvalign) - Vertical alignment

#### Example

```js
// Range object
let range = Application.Range('A1')

// Set alignment: bottom
range.VerticalAlignment = Application.Enum.XlVAlign.xlVAlignBottom

// Get alignment
console.log(range.VerticalAlignment) //-4107 (bottom)
```

## IndentLevel

Cell indent

#### Data type

Number - Indent level

#### Example

```js
// Range object
let range = Application.Range('A1')

// Write value to cell
console.log(range.IndentLevel) // 0
range.IndentLevel = 10
console.log(range.IndentLevel) // 10
```

## WrapText

Get or set wrap text for the range

#### Data type

Boolean - Wrap text

#### Example

```js
// Range object
let range = Application.Range('A1')
// Get wrap text
console.log('Wrap text:', range.WrapText) // true
range.WrapText = false
console.log('Wrap text:', range.WrapText) // false
```

## MergeArea

Merged range containing the cell

#### Data type

Range - Merged area as a Range object

#### Example

```js
// Range object
let range = Application.Range('A1:D2')
// Merge cells.
range.Merge()
// Set merged area value to merge
range.MergeArea.Value = 'merge'
```

## MergeCells

Whether the range contains merged cells

#### Data type

Boolean - Contains merged cells

#### Example

```js
// Range object
let range = Application.Range('A1:D2')
// Merge cells.
range.Merge()
// Whether the range contains merged cells
console.log(range.MergeCells) // true
```

## Cells

All cells in the range as a Range (use Range members).

#### Data type

[Range](/docs/en-US/api/excel/workbook/Range.html) - Cells collection in the range

#### Example

```js
// Range object
let range = Application.Range('A1:B2')

// Cells collection
let cells = range.Cells

// First cell
let cell = cells.Item(1)
console.log(cell.Address()) // $A$1
```

## Rows

Rows in the range as a Range.

#### Data type

[Range](/docs/en-US/api/excel/workbook/Range.html) - Rows in the range

#### Example

```js
// A1:B2 spans 2 rows
console.log(Application.Range('A1:B2').Rows.Count) // 2

// A1:D4 spans 4 rows
console.log(Application.Range('A1:D4').Rows.Count) // 4
```

## Columns

Columns in the range as a Range.

#### Data type

[Range](/docs/en-US/api/excel/workbook/Range.html) - Columns in the range

#### Example

```js
// A1:B2 spans 2 columns
console.log(Application.Range('A1:B2').Columns.Count) // 2

// A1:D4 spans 4 columns
console.log(Application.Range('A1:D4').Columns.Count) // 4
```

## EntireRow

Full rows intersecting the range

#### Data type

[Range](/docs/en-US/api/excel/workbook/Range.html) - Full rows intersecting the range

#### Example

```js
// One full row has 16384 cells
console.log(Application.Range('A1').EntireRow.Count) // 16384

// Two rows
console.log(Application.Range('A1:B2').EntireRow.Count) // 32768
```

## EntireColumn

Full columns intersecting the range

#### Data type

[Range](/docs/en-US/api/excel/workbook/Range.html) - Full columns intersecting the range

#### Example

```js
// One full column has 1048576 cells
console.log(Application.Range('A1').EntireColumn.Count) // 1048576

// Two columns
console.log(Application.Range('A1:B2').EntireColumn.Count) // 2097152
```

## Row

First row number in the range

#### Data type

Number - First row number in the range

#### Example

```js
let range = Application.Range('A1:D4')
// First row of A1:D4 is 1
console.log(range.Row) // 1
```

## RowEnd

Last row number in the range

#### Data type

Number - Last row number in the range

#### Example

```js
let range = Application.Range('A1:D4')
// Last row of A1:D4 is 4
console.log(range.RowEnd) // 4
```

## Column

First column number in the range

#### Data type

Number - First column number in the range

#### Example

```js
let range = Application.Range('A1:D4')
// Left column of A1:D4 is column 1
console.log(range.Column) // 1
```

## ColumnEnd

Last column number in the range

#### Data type

Number - Last column number in the range

#### Example

```js
let range = Application.Range('A1:D4')
// Right column of A1:D4 is column 4
console.log(range.ColumnEnd) // 4
```

## Borders

Borders collection

#### Data type

[Border](/docs/en-US/api/excel/workbook/Border.html) collection - Border objects

#### Example

```js
// Borders collection
let borders = Application.Range('A1').Borders

// Set A1 diagonal-down border color to red
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FF0000'
```

### Borders.Item()

A single border in the Borders collection for a range or style

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Index | Enum |  | Yes | Border to retrieve; see [Enum.XlBordersIndex](/docs/en-US/api/excel/workbook/Enum.html#xlbordersindex) |

#### Return type

[Border](#border) - A single border in the Borders collection for a range or style

#### Example

```js
// Borders collection
let borders = Application.Range('A1').Borders

// Set A1 diagonal-down border color to yellow
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FFFF00'
```

## BorderAround()

Add a border and set Color, LineStyle, and Weight.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| LineStyle | [LineStyle](/docs/en-US/api/excel/workbook/Enum.html#xllinestyle) | null | No | XlLineStyle constant for line style |
| BorderWeight | [BorderWeight](/docs/en-US/api/excel/workbook/Enum.html#xlborderweight) | null | No | Border weight |
| Color | string | null | No | Border color as RGB, e.g. `#ff0000` |

#### Example

```js
// Red dashed border on B2
Range('B2').BorderAround(xlDash, xlHairline, '#ff0000')
// Green dash-dot border on C3
Range('C3').BorderAround(xlDashDot, xlHairline, '#00ff00')
```

## Each()

Iterate each cell; avoid inserting or deleting rows/columns during iteration for predictable results.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| callback | Function | null | Yes | Like Array.prototype.forEach |

#### Example

```js
// Range object
let range = Application.Range('A1:D2')

// Visit each cell
range.Each(function (item) {
  // Print cell address
  console.log(item.Address()) // $A$1...$D$2
})
```

## Item()

Returns a Range for the given index within the range.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| RowIndex | Number |  | Yes | Relative row within the range if ColumnIndex is provided; otherwise index into the subrange |
| ColumnIndex | Number |  | No | Relative column within the range |

#### Return type

Range - Cell at the given position

#### Example

```js
// Range object
const range = Application.Range('A1:D2')
// Sub-item example: B2
console.log(range.Item(2, 2).Address()) //$B$2
```

## Offset()

Offsets the range and returns the new Range.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| RowOffset | Number | 0 | No | Row offset; positive down, negative up |
| ColumnOffset | Number | 0 | No | Column offset; positive right, negative left |

#### Example

```js
// Range object
let range = Application.Range('A1:D2')
// Sub-item example: B2
console.log(range.Address()) //$A$1:$D$2
// Offset the range
let newRange = range.Offset(2, 2)
// Sub-item example: B2
console.log(newRange.Address()) //$C$3:$F$4
```

## Replace()

Replace text inside cells.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| What | String | (none) | Yes | Text to find |
| Replacement | String | (none) | Yes | Replacement text |
| LookAt | Enum | xlWhole | No | XlLookAt: xlWhole or xlPart; see examples |
| SearchOrder | Enum | xlByRows | No | XlSearchOrder: xlByRows or xlByColumns; see examples |
| MatchCase | Boolean | false | No | If true, match case |
| MatchByte | Boolean | false | No | If true, distinguish full-width and half-width characters |

#### Example

```js
// Range object
let range = Application.Range('A1:D2')
// Replace all occurrences of XiaoShan with DaShan (partial match)
range.Replace('foo', 'bar', xlPart)
// Replace whole cell when value equals exactly
range.Replace('foo', 'bar', xlWhole)
// Case sensitive
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, true)
// Case insensitive
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, false)
// Byte-width sensitivity on
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, true, true)
// Byte-width sensitivity off
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, true, false)
```

## Delete()

Delete cells, rows, or columns. Deleting cells shifts remaining cells left.

#### Example: Delete rows

```js
// Range object
let range = Application.Range('B5:D10')
// Delete rows 5-10 covered by this range
range.EntireRow.Delete()
```

#### Example: Delete columns

```js
let range = Application.Range('B5:D10')
// Delete columns 2-4 covered by this range
range.EntireColumn.Delete()
```

## Insert()

Insert cells, rows, or columns.

#### Example: Insert rows

```js
let range = Application.Range('B5:D10')
// Insert a row above row 5; existing row 5 shifts down
range.EntireRow.Insert()
```

#### Example: Insert columns

```js
let range = Application.Range('B5:D10')
// Insert a column left of column 2; existing column 2 shifts right
range.EntireColumn.Insert()
```

## InsertImage()

Insert a picture into a cell.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| dataURL | string | undefined | Yes | Image as a Base64 data URL string |

#### Example

```js
// Cell E1
const range = Range('E1')
// Insert image into target cell
range.InsertImage(
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII='
)
```

## Merge()

Merge cells.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Across | boolean | false | Yes | If true, merge each row in the range into one cell across columns. |

#### Example

```js
const range = Application.Range('A1:D2')
// Merge cells.
range.Merge()
```

## UnMerge()

Unmerge cells.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| CancelCenter | Boolean | false | No | Whether to unmerge and clear center alignment |

#### Example

```js
const range = Application.Range('A1:D2')

// Unmerge
range.UnMerge()
```

## Address()

Address string using macro reference style.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| RowAbsolute | Boolean |  | No | If true, return the row part as an absolute reference. Default true |
| ColumnAbsolute | Boolean |  | No | If true, return the column part as an absolute reference. Default true |
| ReferenceStyle | Enum |  | No | Reference style; default xlA1. See [Enum.XlReferenceStyle](/docs/en-US/api/excel/workbook/Enum.html#xlreferencestyle) |
| External | Boolean |  | No | If true, return an external reference; if false, local reference. Default false |
| RelativeTo | Range |  | No | When both row and column are relative and style is xlR1C1, anchor range for the relative address |

#### Return type

String - Address string in macro reference style

#### Example

```js
// Range object
let range = Application.Range('A1')

// Address string
console.log('address1:', range.Address())

console.log('address2:', range.Address(false, false))

console.log(
  'address3:',
  range.Address(true, true, Application.Enum.XlReferenceStyle.xlR1C1)
)
```

## AddComment()

Add a comment.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Text | String |  | No | Comment text |

#### Example

```js
let range = Application.Range('A1')
// Add comment to A1
range.AddComment('WebOffice')
```

## ClearComments()

Clear comments in the range.

#### Example

```js
let range = Application.Range('A1')
// Clear comments in the range.
range.ClearComments()
```

## Clear()

Clear values and formats.

#### Example

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.Clear()
console.log(range.Value) //[["",""],["",""]]
```

## ClearContents()

Clear contents.

#### Example

```js
let range = Application.Range('A1')
// Clear contents.
range.ClearContents()
```

## ClearFormats()

Clear formats.

#### Example

```js
let range = Application.Range('A1')
// Clear formats.
range.ClearFormats()
```

## ClearHyperlinks()

Clear hyperlink formatting.

#### Example

```js
let range = Application.Range('A1')
// Clear hyperlink formatting.
range.ClearHyperlinks()
```

## Contain()

Whether two ranges overlap.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Range | Range |  | Yes | Another Range object |

#### Return type

Boolean - Whether ranges overlap

#### Example

```js
// Range object
let range = Application.Range('A1:D2')

// Second range object
let newRange = Application.Range('A1:B4')

// Test overlap
console.log(range.Contain(newRange))
```

## Copy()

Copy the range to the clipboard.

#### Return type

Boolean - Whether copy succeeded

#### Example

```js
// Copy A1
Range('A1').Copy()
// Paste to B1
Range('B1').PasteSpecial()
```

## Cut()

Cut the range to a destination.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Destination | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Destination range |

#### Return type

Boolean - Whether copy succeeded

#### Example

```js
// Move A1 to A2
Range('A1').Cut(Range('A2'))
```

## PasteSpecial()

Pastes clipboard contents with options such as values, formats, or formulas.

Note

Before PasteSpecial(), copy a range to the clipboard; pasting with an empty clipboard throws an error.

Use [Copy()](#copy) first, then PasteSpecial() with the desired options.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Paste | [XlPasteType](/docs/en-US/api/excel/workbook/Enum.html#xlpastetype) | xlPasteAll | No | Paste type, e.g. xlPasteAll or xlPasteValues |
| Operation | [XlPasteSpecialOperation](/docs/en-US/api/excel/workbook/Enum.html#xlpastespecialoperation) | xlPasteSpecialOperationNone | No | Paste operation, e.g. xlPasteSpecialOperationAdd |
| SkipBlanks | Boolean | false | No | If true, skip blank cells from the source when pasting |
| Transpose | Boolean | false | No | If true, transpose rows and columns when pasting |

#### Return type

Undefined

#### Example 1

Paste values only, e.g. paste B1 into A1 as values.

```js
Range('B1').Copy()
Range('A1').PasteSpecial(xlPasteValues)
```

#### Example 2

Paste formats from A1:B2 to C1:D2.

```js
Range('A1:B2').Copy()
Range('C1:D2').PasteSpecial(xlPasteFormats)
```

#### Example 3

Paste values and formats from A1:B2 to A3:B4.

```js
Range('A1:B2').Copy()
Range('A3:B4').PasteSpecial(xlPasteAll)
```

#### Example 4

Transpose paste A1:B2 to C1:D2 (rows and columns swap).

```js
Range('A1:B2').Copy()
Range('C1:D2').PasteSpecial(
  xlPasteValues,
  xlPasteSpecialOperationNone,
  false,
  true
)
```

## FillLeft()

Fill left using the rightmost column's values.

#### Example

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillLeft()
console.log(range.Value) // [["B1","B1"],["B2","B2"]]
```

## FillRight()

Fill right using the leftmost column's values.

#### Example

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillRight()
console.log(range.Value) // [["A1","A1"],["A2","A2"]]
```

## FillDown()

Fill down using the top row's values.

#### Example

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillDown()
console.log(range.Value) // [["A1","B1"],["A1","B1"]]
```

## FillUp()

Fill up using the bottom row's values.

#### Example

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillUp()
console.log(range.Value) //[["A2","B2"],["A2","B2"]]
```

## AutoFill()

AutoFill the destination range.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Destination | Range |  | No | Destination range; must include the source range |
| Type | Enum | Enum.XlAutoFillType.xlFillDefault | No | Fill type; see [Enum.XlAutoFillType](/docs/en-US/api/excel/workbook/Enum.html#xlautofilltype) |

#### Example

```js
// Range object
let range = Application.Range('A1')
range.Value = 1
// Cells to fill
let fillRange = Application.Range('A1:A20')
// AutoFill the destination range.
range.AutoFill(fillRange)
fillRange.Each(item => {
  console.log(item.Text()) // AutoFilled: 1,2,3,4...20
})
```

## AutoFilter()

Apply AutoFilter to the range.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Field | Variant | null | No | Field index from the left of the list; the leftmost field is 1 |
| Criteria1 | Variant | null | No | First criterion. Use "=" for blanks, "<>" for non-blanks. If omitted, all values. For xlTop10Items, specifies the count. |
| Operator | [XlAutoFilterOperator](/docs/en-US/api/excel/workbook/Enum.html#xlautofilteroperator) | null | No | Filter operator; see [XlAutoFilterOperator](/docs/en-US/api/excel/workbook/Enum.html#xlautofilteroperator) |
| Criteria2 | Variant | null | No | Second criterion with Criteria1 and Operator. Also used for date filters as Array(Level, Date) where Level is 0-2 (year, month, day) and Date is a valid date in the range |
| SubField | Variant | null | No | Field within a structured type (e.g. Population or Volume). Omit for display values |
| VisibleDropDown | Variant | true | No | If `true`, show the AutoFilter dropdown for the field; if `false`, hide it. Default `true` |

#### Example 1

Show the bottom 11 rows by values in column B.

```js
// When no range is passed, UsedRange is used as the filter region
const filterRange = ActiveSheet.UsedRange
// Column B
let filterField = 2
// Data column for the filter
const filterColumnRange = filterRange.Columns(filterField)
// Row count minus header
const filterRangeCount = filterColumnRange.Rows.Count - 1
const bottom11Value = 11
const bottomNum =
  filterRangeCount < bottom11Value ? filterRangeCount : bottom11Value
let bottom11 = WorksheetFunction.Small(filterColumnRange, bottomNum)
filterRange.AutoFilter(filterField, '<=' + bottom11)
```

#### Example 2

When column 1 is "1 Class", fill column 5 cells red if value < 60 and green if value > 90

```js
// Use used range as filter area
const filterRange = ActiveSheet.UsedRange
// Filter column 1
const filterColumn1 = 1
// Filter value for column 1
let filterValue = '1 Class'
// Apply filter
filterRange.AutoFilter(filterColumn1, filterValue)
// Filter column 5
const filterColumn2 = 5
// Criterion less than 60
filterValue = '<60'
// Apply filter
filterRange.AutoFilter(filterColumn2, filterValue)
// Data without header row
const dataRange = ActiveSheet.UsedRange.Offset(1).Resize(
  ActiveSheet.UsedRange.Rows.Count - 1
)
// Column range
const columnRange = dataRange.Columns(filterColumn2)
// Visible cells after filter
const visibleRange = columnRange.SpecialCells(xlCellTypeVisible)
// Define color red
const redColor = RGB(255, 0, 0)
// Set color to red
visibleRange.Interior.Color = redColor
// Clear filter on column 5
filterRange.AutoFilter(filterColumn2)
// Criterion greater than 90
const filterValue2 = '>90'
// Apply filter
filterRange.AutoFilter(filterColumn2, filterValue2)
// Visible cells after filter
const visibleRange2 = columnRange.SpecialCells(xlCellTypeVisible)
// Define color green
const greenColor = RGB(0, 255, 0)
// Set color to green
visibleRange2.Interior.Color = greenColor
// Clear filter on column 1
filterRange.AutoFilter(filterColumn1)
// Clear filter on column 5
filterRange.AutoFilter(filterColumn2)
```

#### Example 3

Filter rows where column A matches the given dates and column C is above 10000

```js
// Default filter range is UsedRange
const filterRange = ActiveSheet.UsedRange
// Filter column A
let filterField = 1
// Date criteria array
let criteria = new Array()
// Date level enum
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// Add day criterion 2023/4/10
criteria.push(FilterDateTimeEnum.DAY, '2023/4/10')
// Add month criterion Dec 2022
criteria.push(FilterDateTimeEnum.MONTH, '2022/12/1')
// Apply filter
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
// Filter column C
filterField = 3
// Criterion > 10000
let criteria2 = '>10000'
// Apply filter
filterRange.AutoFilter(filterField, criteria2)
```

#### Example 4

Filter column A for the listed year/month/day combinations

```js
// Default filter range is UsedRange
const filterRange = ActiveSheet.UsedRange
// Filter column A
let filterField = 1
// Date criteria array
let criteria = new Array()
// Date level enum
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// Add day criterion 2023/4/10
criteria.push(FilterDateTimeEnum.YEAR, '2023/1/1')
// Add Mar 2022
criteria.push(FilterDateTimeEnum.MONTH, '2022/3/1')
// Add Apr 2022
criteria.push(FilterDateTimeEnum.MONTH, '2022/4/1')
// Add 2022-05-01
criteria.push(FilterDateTimeEnum.DAY, '2022/5/1')
// Add 2022-05-02
criteria.push(FilterDateTimeEnum.DAY, '2022/5/2')
// Apply filter
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
```

#### Example 5

Show rows where D is this quarter, C is year-to-date, and A is Product A

```js
// Default filter range is UsedRange
const filterRange = ActiveSheet.UsedRange
// Filter column D
let filterField = 4
// Quarter filter (this/last/next quarter)
let criteria = xlFilterThisQuarter
// Apply filter
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// Filter column C
filterField = 3
// Year-to-date filter
criteria = xlFilterYearToDate
// Apply filter
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// Filter column A
filterField = 1
// Value Product A
criteria = 'ProductA'
// Apply filter
filterRange.AutoFilter(filterField, criteria)
```

#### Example 6

Filter column 2 contains Zhang or starts with Li

```js
// Default filter range is UsedRange
const filterRange = ActiveSheet.UsedRange
// Filter column 2
let filterField = 2
// Contains Zhang
let criteria = '=*Zhang*'
// Starts with Li
let criteria2 = '=Li*'
// Apply with OR
filterRange.AutoFilter(filterField, criteria, xlOr, criteria2)
```

#### Example 7

Show top 5 rows by values in column A

```js
// Default filter range is UsedRange
const filterRange = ActiveSheet.UsedRange
// Filter column A
let filterField = 1
// Data column range
const filterColumnRange = filterRange.Columns(filterField)
// Row count minus header
const filterRangeCount = filterColumnRange.Rows.Count - 1
// Get top 5 values
const top5Value = 5
// Clamp to available rows
const topNum = filterRangeCount < top5Value ? filterRangeCount : top5Value
// Fifth largest value
let top5 = WorksheetFunction.Large(filterColumnRange, topNum)
// Apply filter
filterRange.AutoFilter(filterField, '>=' + top5)
```

#### Example 8

For visible rows where column C is Zhang, set column F to KaiTi font

```js
// Use used range as filter area
const filterRange = ActiveSheet.UsedRange
// Filter column 3
const filterColumn = 3
// Value Zhang
const filterValue = ['Zhang']
// Apply filter
filterRange.AutoFilter(filterColumn, filterValue)

// Get filter result
const filteredRange =
  ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible)

// Visible cells in column F
const filteredRangeFColumn =
  ActiveSheet.AutoFilter.Range.Columns('F').SpecialCells(xlCellTypeVisible)

// Apply KaiTi to column F
filteredRangeFColumn.Font.Name = 'KaiTi'
```

#### Example 9

Show rows where column 2 is Li and column 4 is last week

```js
// Default filter range is UsedRange
const filterRange = ActiveSheet.UsedRange
// Filter column 2
let filterField = 2
// Value Li
let criteria = 'Li'
// Apply filter
filterRange.AutoFilter(filterField, criteria)
// Filter column 4
filterField = 4
// Last week filter
criteria = xlFilterLastWeek
// Apply filter
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
```

## AutoFit()

Autofit column widths or row heights.

Must be called on [Columns](#column) or [Rows](#rows); otherwise it has no effect.

One column-width unit equals the width of one character in the Normal style.

#### Example

```js
// AAutofit column width
const range = Application.Range('A1')
range.Columns.AutoFit()

// Autofit row 1 height
const range = Application.Range('A1')
range.Rows.AutoFit()
```

## Select()

Select the range; updates Application.Selection.

#### Example

```js
// Range object
let range = Application.Range('A1')
range.Select()
let selection = Application.Selection
console.log(selection.Address())
```

## TextToColumns()

Split cell text into multiple columns using delimiters.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Destination | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | No | Destination for split columns; if omitted, overwrites source |
| DataType | [XlTextParsingType](/docs/en-US/api/excel/workbook/Enum.html#xltextparsingtype) | xlDelimited | No | Text format to parse |
| TextQualifier | [XlTextQualifier](/docs/en-US/api/excel/workbook/Enum.html#xltextqualifier) | xlTextQualifierDoubleQuote | No | Text delimiter |
| ConsecutiveDelimiter | Boolean | false | No | Treat consecutive delimiters as one |
| Tab | Boolean | false | No | Use tab as delimiter |
| Semicolon | Boolean | false | No | Use semicolon as delimiter |
| Comma | Boolean | false | No | Use comma as delimiter |
| Space | Boolean | false | No | Use space as delimiter |
| Other | Boolean | false | No | Use `OtherChar` as delimiter |
| OtherChar | String | `''` | No | Custom delimiter character; default empty |
| FieldInfo | Array | null | No | Per-column parse options array |
| DecimalSeparator | String | System settings | No | Decimal separator when parsing numbers |
| ThousandsSeparator | String | System settings | No | Thousands separator when parsing numbers |
| TrailingMinusNumbers | String | null | No | Numbers that start with a minus sign |

Tip

About `FieldInfo`: this array describes how each column is parsed, depending on `DataType`.

For delimited data, each entry is a two-element array: column index (1-based) and a [XlColumnDataType](/docs/en-US/api/excel/workbook/Enum.html#xlcolumndatatype) value. Columns without an entry use General. For fixed-width data, the first element is the starting character index.

#### Return type

undefined

#### Example 1

Split A1:A3 on tabs, column B on semicolons, C1 on commas, and D1 on spaces.

```js
// 1. Tab-delimited A1:A3
Range('A1:A3').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined
)
// 2. Semicolon-delimited column B
Range('B:B').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined
)
// 3. Comma-delimited C1
Range('C1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined
)
// 4. Space-delimited D1
Range('D1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined
)
```

#### Example 2

Split A1 on `@`, A2 on `#`, and A3 on `*`, using OtherChar.

```js
// 1. A1 with OtherChar '@'
Range('A1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '@',
  undefined
)
// 2. A2 with OtherChar '#'
Range('A2').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '#',
  undefined
)
// 3. A3 with OtherChar '*'
Range('A3').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '*',
  undefined
)
```

#### Example 3

Split A1:A10 on space and semicolon into B1; split column B on comma, tab, and `=` into column D; split column C on tab, comma, semicolon, and space.

```js
// 1. Space + semicolon into B1
Range('A1:A10').TextToColumns(
  Range('B1'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  true,
  undefined,
  undefined,
  undefined
)
// 2. Comma, tab, and '=' on column B into column D
Range('B:B').TextToColumns(
  Range('D:D'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  undefined,
  true,
  undefined,
  true,
  '=',
  undefined
)
// 3. Tab, comma, semicolon, and space on column C
Range('C:C').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  true,
  true,
  true,
  undefined,
  undefined,
  undefined
)
```

#### Example 4

Comma-split A1 treating consecutive delimiters as one; split B1:B5 on `&` and tab with consecutive delimiters merged; split column C on `!` treating consecutive delimiters separately.

```js
// 1. Comma, ConsecutiveDelimiter true
Range('A1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  true,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined
)
// 2. OtherChar '&' plus tab, ConsecutiveDelimiter true
Range('B1:B5').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  true,
  undefined,
  undefined,
  undefined,
  true,
  true,
  '&',
  undefined
)
// 3. OtherChar '!', ConsecutiveDelimiter false
Range('C:C').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  false,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '!',
  undefined
)
```

#### Example 5

Split A1 on spaces into B1 with FieldInfo column types general, text, and YMD date. Split C1 on semicolons into C3 with DMY, MYD, DYM, and YDM date columns.

```js
// 1. Space-delimited with FieldInfo [column, xlColumnDataType]
Range('A1').TextToColumns(
  Range('B1'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  [
    [1, xlGeneralFormat],
    [2, xlTextFormat],
    [3, xlYMDFormat]
  ]
)
// 2. Semicolon-delimited with per-column date formats
Range('C1').TextToColumns(
  Range('C3'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  [
    [2, xlDMYFormat],
    [3, xlMYDFormat],
    [4, xlDYMFormat],
    [5, xlYDMFormat]
  ]
)
```
