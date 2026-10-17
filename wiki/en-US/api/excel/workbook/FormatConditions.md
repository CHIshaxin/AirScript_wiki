---
title: "FormatConditions"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/FormatConditions.html
updated: 1789700916000
hash: DaYds8OW
---

# FormatConditions

The FormatConditions collection controls conditional formatting in Excel.

Conditional formatting applies formats based on cell values, formulas, or other rules so data is easier to read.

#### Properties

| Property | Data type | Description |
| --- | --- | --- |
| [Count](#count) | Number | Number of objects in the FormatConditions collection |

#### Methods

| Method | Return type | Description |
| --- | --- | --- |
| [Add()](#add) | [FormatCondition](FormatCondition.md) | Adds a conditional format to the collection |
| [AddAboveAverage()](#addaboveaverage) | AboveAverage | New AboveAverage rule for the range |
| [AddIconSetCondition()](#addiconsetcondition) | IconSetCondition | Icon set conditional format for the range |
| [AddColorScale()](#addcolorscale) | ColorScale | Color scale rule showing relative cell values with a gradient |
| [AddTop10()](#addtop10) | Top10 | Highlights highest or lowest values by rank or percentage |
| [AddUniqueValues()](#adduniquevalues) | UniqueValues | New UniqueValues rule for the range |
| [Delete()](#delete) | undefined | Deletes conditional formats on the range |
| [Item()](#item) | [FormatCondition](FormatCondition.md) | Returns one conditional format from the collection |

## Count

Number of objects in the FormatConditions collection.

#### Data type

Number - Object count

#### Example

```js
const formatConditions = Range('A:A').FormatConditions
const count = formatConditions.Count

console.log(count)
```

## Add()

Adds a conditional format to the collection.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Type | [XlFormatConditionType](Enum.md#xlformatconditiontype) | null | Yes | Condition type |
| Operator | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | null | Yes | Operator; ignored if Type is `xlExpression` |
| Formula1 | String | null | Yes | Value or expression: constant, string, reference, or formula |
| Formula2 | String | null | No | Second part when Operator is `xlBetween` or `xlNotBetween` (otherwise ignored). Constant, string, reference, or formula |

#### Return value

[FormatCondition](FormatCondition.md) - The added conditional format

#### Example

```js
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
const formula1 = `=AND(A9=2, B9>${ave})`
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
formatCondition1.SetFirstPriority()
formatCondition1.Font.Strikethrough = true
```

## AddAboveAverage()

Returns a new AboveAverage object for the range.

AboveAverage highlights values above or below the average or standard deviation, for example in performance reviews.

#### Return value

AboveAverage object

#### Example

```js
const formatconditions = Range('D:D').FormatConditions
const aboveAverageCondition = formatconditions.AddAboveAverage()
aboveAverageCondition.AboveBelow = xlAboveAverage
aboveAverageCondition.SetFirstPriority()
aboveAverageCondition.Font.Color = RGB(255, 255, 0)
```

## AddIconSetCondition()

Returns a new IconSetCondition for an icon set rule on the range.

Icon sets annotate data and split it into three to five bands by threshold.

#### Return value

IconSetCondition object

#### Example

```js
// Icon set on columns D and F: >=130, between 90 and 130, <=90

const bIgnoreHeader = true
const targetWorkbook = ActiveWorkbook
const targetWorksheet = ActiveSheet

const targetColumns = ['D', 'F']
for (let i = 0; i < targetColumns.length; ++i) {
  const targetColumn = targetColumns[i]
  const beginRow = targetWorksheet.UsedRange.Row + (bIgnoreHeader ? 1 : 0)
  const endRow =
    targetWorksheet.UsedRange.Row + targetWorksheet.UsedRange.Rows.Count - 1
  const targetRangeStr = targetColumn + beginRow + ':' + targetColumn + endRow
  let targetRange = targetWorksheet.Range(targetRangeStr)
  targetRange.Select()
  let iconSetCondition = targetRange.FormatConditions.AddIconSetCondition()
  iconSetCondition.SetFirstPriority()
  iconSetCondition.IconSet = targetWorkbook.IconSets.Item(xl3TrafficLights1)

  const maxIndex = 3
  const middleIndex = 2
  let IconCriterion3 = iconSetCondition.IconCriteria.Item(maxIndex)
  IconCriterion3.Type = xlConditionValueNumber
  IconCriterion3.Operator = xlGreaterEqual
  IconCriterion3.Value = 130

  let IconCriterion2 = iconSetCondition.IconCriteria.Item(middleIndex)
  IconCriterion2.Type = xlConditionValueNumber
  IconCriterion2.Operator = xlGreater
  IconCriterion2.Value = 90
}
```

## AddColorScale()

Returns a new ColorScale object: a color gradient showing relative cell values.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| ColorScaleType | Number | null | Yes | 2 for two-color scale, 3 for three-color scale |

#### Return value

ColorScale object

#### Example

```js
// Two-color scale on column B: min red, max green

const formatConditions = Range('B:B').FormatConditions
const color2ScaleCondition = formatConditions.AddColorScale(2)
color2ScaleCondition.SetFirstPriority()
const colorScaleCriteria1 = color2ScaleCondition.ColorScaleCriteria.Item(1)
colorScaleCriteria1.Type = xlConditionValueLowestValue
colorScaleCriteria1.FormatColor.Color = RGB(255, 0, 0)
const colorScaleCriteria2 = color2ScaleCondition.ColorScaleCriteria.Item(2)
colorScaleCriteria2.Type = xlConditionValueHighestValue
colorScaleCriteria2.FormatColor.Color = RGB(0, 255, 0)
```

## AddTop10()

Returns a Top10 object for top or bottom values in the range.

Use Top10 to highlight the highest or lowest values by rank or percentage.

#### Return value

Top10 object

#### Example

```js
// Bottom 5 values in column L in green font

const bIgnoreHeader = true
const targetWorksheet = ActiveSheet
const targetColumn = 'L'
const beginRow = targetWorksheet.UsedRange.Row + (bIgnoreHeader ? 1 : 0)
const endRow =
  targetWorksheet.UsedRange.Row + targetWorksheet.UsedRange.Rows.Count - 1
const targetRangeStr = targetColumn + beginRow + ':' + targetColumn + endRow
let targetRange = targetWorksheet.Range(targetRangeStr)
targetRange.Select()

let top10 = targetRange.FormatConditions.AddTop10()
top10.Percent = false
top10.Rank = 5
top10.TopBottom = xlTop10Bottom
const greenColor = RGB(0, 255, 0)
top10.Font.Color = greenColor
top10.Font.TintAndShade = 0
top10.StopIfTrue = false
```

## AddUniqueValues()

Returns a new UniqueValues object for the range.

Use UniqueValues to highlight unique or duplicate cells.

#### Return value

UniqueValues object

#### Example

```js
// Highlight duplicates in B2:G18 with red fill

const range = Range('B2:G18')
range.FormatConditions.Delete()
let formatCondition = range.FormatConditions.AddUniqueValues()
formatCondition.DupeUnique = xlDuplicate
const redColor = RGB(255, 0, 0)
formatCondition.Interior.Color = redColor
```

## Delete()

Deletes conditional formats on the range.

#### Return value

undefined

```js
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Delete()
```

## Item()

Returns one conditional format from the collection.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Index | Number / String | null | Yes | One-based index in the collection |

#### Return value

[FormatCondition](FormatCondition.md) object

#### Example

```js
const formatConditions = Range('B:B').FormatConditions
const destRange = Range('A:A')
const count = formatConditions.Count
for (let i = 1; i <= count; i++) {
  formatConditions.Item(i).ModifyAppliesToRange(destRange)
}
```
