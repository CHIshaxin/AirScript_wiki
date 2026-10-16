---
title: "Sort"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Sort.html
updated: 1789700916000
hash: WZAJyRaX
---

# Sort

Sort object; set the target range, then apply sorting.

#### Properties

| Property | Data type | Description |
|---|---|---|
| [Header](#header) | [XlYesNoGuess](/docs/en-US/api/excel/workbook/Enum.html#xlyesnoguess) | Whether the first row contains headers |
| [MatchCase](#matchcase) | Boolean | Case-sensitive sort |
| [Orientation](#orientation) | [XlSortOrientation](/docs/en-US/api/excel/workbook/Enum.html#xlsortorientation) | Sort orientation |
| [Rng](#rng) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Range containing values to be sorted |
| [SortFields](#sortfields) | [SortFields](/docs/en-US/api/excel/workbook/SortFields.html) | Collection of sort fields for this **Sort** object |
| [SortMethod](#sortmethod) | [XlSortMethod](/docs/en-US/api/excel/workbook/Enum.html#xlsortmethod) | East Asian sort method (phonetic or stroke) |

#### Methods

| Method | Return type | Description |
|---|---|---|
| [Apply()](#apply) | undefined | Sorts the range using the current sort state |
| [SetRange()](#setrange) | undefined | Sets the range to be sorted |

#### Usage examples

Common sort scenarios on a worksheet:

Example 1. Sort column C descending and column D ascending

```js
// Current used range
const range = ActiveSheet.UsedRange
// Sort object
const sort = ActiveSheet.Sort
// Sort fields
const sortFields = sort.SortFields
// Clear previous keys
sortFields.Clear()
// Column C by value descending
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
// Column D by value ascending
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
// Header: xlGuess auto, xlYes has header, xlNo no header
sort.Header = xlGuess
// Case sensitivity
sort.MatchCase = false
// East Asian sort: xlPinYin or xlStroke
sort.SortMethod = xlPinYin
// xlSortColumns vs xlSortRows
sort.Orientation = xlSortColumns
// SetRange required before Apply
sort.SetRange(range)
sort.Apply()
```

Example 2. Put red cell color in column G on top; red in column F at bottom

```js
// Current used range
const range = ActiveSheet.UsedRange
// Sort object
const sort = ActiveSheet.Sort
// Sort fields
const sortFields = sort.SortFields
// Clear previous keys
sortFields.Clear()
// Column G: red on top
// Args: key range, sort-on (values/font color/cell color), order (asc/desc)
const sortField1 = sortFields.Add(
  Range('G:G').Item(1, 1),
  xlSortOnCellColor,
  xlAscending
)
// Sort color red
sortField1.SortOnValue.Color = RGB(255, 0, 0)
// Column F: red at bottom
const sortField2 = sortFields.Add(
  Range('F:F').Item(1, 1),
  xlSortOnCellColor,
  xlDescending
)
// Sort color red
sortField1.SortOnValue.Color = RGB(255, 0, 0)
sort.Header = xlGuess
sort.MatchCase = false
sort.SortMethod = xlPinYin
sort.Orientation = xlSortColumns
sort.SetRange(range)
sort.Apply()
```

Example 3. Sheet has columns A through H; sort so smaller values in column B come first

```js
// Sort object
const sort = ActiveSheet.Sort

const keyColumn = 'B'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

const sortFields = sort.SortFields
const sortKeyRange = ActiveSheet.Range(
  keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow
)
sortFields.Clear()
// Sort by values ascending
sortFields.Add(sortKeyRange, xlSortOnValues, xlAscending)

// Range to sort
const applyBeginRow = ActiveSheet.UsedRange.Row
const applyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const applyRange = ActiveSheet.Range(`A${applyBeginRow}:H${applyEndRow}`)
// SetRange before Apply
sort.SetRange(applyRange)
sort.Apply()
```

Example 4. Sheet has columns A through H; sort column A by stroke count descending

```js
// Sort object
const sort = ActiveSheet.Sort

// Key range
const keyColumn = 'A'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow

// Descending by value
const sortFields = sort.SortFields
sortFields.Clear()
sortFields.Add(
  ActiveSheet.Range(keyRangeStr),
  xlSortOnValues,
  xlDescending,
  '',
  undefined
)
sort.header = xlYes
sort.Orientation = xlSortColumns
// Stroke sort order
sort.SortMethod = xlStroke

// Sort range
const beginRow = ActiveSheet.UsedRange.Row
const endRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

sort.SetRange(ActiveSheet.Range(`A${beginRow}:H${endRow}`))
sort.Apply()
```

Example 5. Sheet has columns A through H; sort by cell color in column B with no fill on top

```js
// Key range
const keyColumn = 'B'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow
let keyRange = ActiveSheet.Range(keyRangeStr)

// Sort by cell color ascending
let sort = ActiveSheet.Sort
let sortFields = sort.SortFields
sortFields.Clear()
sortFields.Add(keyRange, xlSortOnCellColor, xlDescending, '', undefined)
sort.header = xlYes
sort.Orientation = xlSortColumns
sort.SortMethod = xlPinYin

const beginRow = ActiveSheet.UsedRange.Row
const endRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

sort.SetRange(ActiveSheet.Range(`A${beginRow}:H${endRow}`))
sort.Apply()
```

## Header

Whether the first row contains header information. Read/write [XlYesNoGuess](/docs/en-US/api/excel/workbook/Enum.html#xlyesnoguess).

Default is xlNo; use xlGuess to let Excel detect headers.

#### Data type

[XlYesNoGuess](/docs/en-US/api/excel/workbook/Enum.html#xlyesnoguess) - Header row handling

## MatchCase

`true` for case-sensitive sort, `false` otherwise. Read/write.

#### Data type

Boolean - Case sensitivity

## Orientation

Sort orientation. Read/write [XlSortOrientation](/docs/en-US/api/excel/workbook/Enum.html#xlsortorientation).

#### Data type

[XlSortOrientation](/docs/en-US/api/excel/workbook/Enum.html#xlsortorientation) - Sort orientation

## Rng

Returns the range of values to sort. Read-only.

#### Data type

[Range](/docs/en-US/api/excel/workbook/Range.html) - Sort range

## SortFields

Collection of sort fields for this **Sort** object. Read-only.

### Count

Returns the number of objects in the collection. Read-only.

#### Data type

Number - Object count

### Add()

Creates a new sort field and returns a SortField object.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Key | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Key range |
| SortOn | [XlSortOn](/docs/en-US/api/excel/workbook/Enum.html#xlsorton) | null | No | Property to sort by |
| Order | [XlSortOrder](/docs/en-US/api/excel/workbook/Enum.html#xlsortorder) | null | No | Sort order |
| CustomOrder | Variant | null | No | Custom list usage |
| DataOption | [XlSortDataOption](/docs/en-US/api/excel/workbook/Enum.html#xlsortdataoption) | null | No | Data option |

#### Return value

[SortField](/docs/en-US/api/excel/workbook/SortField.html) - Sort field for this **Sort** object

### Clear()

Clears all SortField objects. Call before adding new keys.

#### Return value

undefined

### Item()

Returns a [SortField](/docs/en-US/api/excel/workbook/SortField.html) object. Read-only.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Index | Number / String | null | Yes | Index starting at 1 |

#### Return value

[SortField](/docs/en-US/api/excel/workbook/SortField.html) - Sort field for this **Sort** object

## SortMethod

East Asian sort method (phonetic or stroke). Read/write [XlSortMethod](/docs/en-US/api/excel/workbook/Enum.html#xlsortmethod).

#### Data type

[XlSortMethod](/docs/en-US/api/excel/workbook/Enum.html#xlsortmethod) - East Asian sort method

## Apply()

Sorts the range using the current sort state.

Note

Call [SetRange](#setrange) before applying sort rules, or sorting has no effect.

#### Return value

undefined

## SetRange()

Sets the range to be sorted.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Range the **Sort** object sorts |

#### Return value

undefined
