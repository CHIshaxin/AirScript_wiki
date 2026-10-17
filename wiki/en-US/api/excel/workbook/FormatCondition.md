---
title: "FormatCondition"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/FormatCondition.html
updated: 1789700916000
hash: CSyfM2V8
---

# FormatCondition

A single conditional format within a range.

See the lists below for the FormatCondition object's properties and methods.

#### Properties

| Property | Data type | Description |
| --- | --- | --- |
| [AppliesTo](#appliesto) | [Range](Range.md) | Range where the format rule applies |
| [Borders](#borders) | [Border](Border.md) | Borders collection |
| [Font](#font) | [Font](Font.md) | Font object |
| [Formula1](#formula1) | String | Value or expression for the condition or validation |
| [Formula2](#formula2) | String | Second part for validation or conditions between two values |
| [Interior](#interior) | Interior | Interior fill of the object |
| [NumberFormat](#numberformat) | String | Number format when the rule is true |
| [Operator](#operator) | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | Conditional format operator |
| [Priority](#priority) | Number | Priority of this rule |
| [Type](#type) | [XlFormatConditionType](Enum.md#xlformatconditiontype) | Condition object type |

#### Methods

| Method | Return type | Description |
| --- | --- | --- |
| [Modify()](#modify) | undefined | Modify an existing rule |
| [ModifyAppliesToRange()](#modifyappliestorange) | undefined | Set the range this rule applies to |
| [SetFirstPriority()](#setfirstpriority) | undefined | Set this rule's priority to 1 (evaluated first) |
| [SetLastPriority()](#setlastpriority) | undefined | Set this rule's priority last (increment others) |

## AppliesTo

Returns the Range where the format rule applies.

#### Data type

[Range](Range.md) - Range object

#### Example

```js
// Average of B where A is 2; highlight B above that average when A is 2
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
const formula1 = `=AND(A9=2, B9>${ave})`
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
const appliesToRange = formatCondition1.AppliesTo
console.log(appliesToRange.Address())
```

## Borders

Returns Borders for the range or style (including conditional format borders).

#### Data type

[Border](Border.md) collection

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
formatCondition1.Borders.Item(xlDiagonalDown).Color = '#FFFF00'
```

## Font

Returns the Font object for the formatted range.

#### Data type

[Font](Font.md) - Font object

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
const font = formatCondition1.Font
console.log(font.Color)
```

## Formula1

Value or expression associated with the condition or validation.

Formula1 may be a constant, string, reference, or formula.

#### Data type

String - Value or expression

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
console.log(formatCondition1.Formula1)
```

## Formula2

Second part for conditions or validation.

Used when Operator is xlBetween or xlNotBetween. May be constant, string, reference, or formula.

#### Data type

String - Value or expression

#### Example

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(
  xlCellValue,
  xlBetween,
  '90',
  '100'
)
console.log(betweenCondition.formula2)
```

## Interior

Returns the Interior object (fill).

#### Data type

Interior - Interior object

#### Example

```js
const range = Range('B2:G18')
range.FormatConditions.Delete()
let formatCondition = range.FormatConditions.AddUniqueValues()
formatCondition.DupeUnique = xlDuplicate
const redColor = RGB(255, 0, 0)
formatCondition.Interior.Color = redColor
```

## NumberFormat

Gets or sets the number format applied when the rule evaluates to True.

Uses the same format codes as the Format Cells Number tab.

You can use built-in formats such as "General" or custom codes.

#### Data type

String - Number format

#### Example

```js
let targetFormulaRange = Range('B:B')
let formatCondition1 = targetFormulaRange.FormatConditions
formatCondition1.NumberFormat = '0.0'
```

## Operator

Returns the conditional format operator.

#### Data type

[XlFormatConditionOperator](Enum.md#xlformatconditionoperator) - Operator enum

#### Example

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
console.log(betweenCondition.Operator)
```

## Priority

Gets or sets rule priority. When multiple rules exist, priority controls evaluation order.

#### Data type

Number - Priority

#### Example

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
console.log(betweenCondition.Priority)
```

## Type

Returns the object type.

One of [XlFormatConditionType](Enum.md#xlformatconditiontype):

- xlCellValue
- xlExpression

#### Data type

[XlFormatConditionType](Enum.md#xlformatconditiontype) - Condition type

#### Example

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
console.log(betweenCondition.Type)
```

## Modify()

Modifies an existing conditional format.

### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Type | [XlFormatConditionType](Enum.md#xlformatconditiontype) | null | Yes | Cell value or expression based |
| Operator | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | null | No | Ignored if Type is xlExpression |
| Formula1 | String | null | No | Value or expression |
| Formula2 | String | null | No | Second value or expression |

### Return value

undefined

#### Example

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
betweenCondition.Modify(xlCellValue, xlBetween, '=90', '=110')
```

## ModifyAppliesToRange()

Sets the cell range this rule applies to.

Must be A1-style and fully on the worksheet that owns the FormatConditions collection. Range operators (colon), intersection (space), and union (comma) are allowed; currency symbols are ignored.

### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Range | [Range](Range.md) | null | Yes | Range this rule applies to |

### Return value

undefined

### Example

```js
const formatConditions = Range('B:B').FormatConditions
const destRange = Range('A:A')
const count = formatConditions.Count
for (let i = 1; i <= count; i++) {
  formatConditions.Item(i).ModifyAppliesToRange(destRange)
}
```

## SetFirstPriority()

Sets this rule's priority to 1 so it is evaluated before all others.

If it was not already first, other rules' priorities are increased by one.

Note

Conditional format priority is evaluated at worksheet scope.

### Return value

undefined

### Example

```js
/**
 * Color red: rows A20:A100 where column A is "Zhang" and dates D20:D100 are in the current month.
 * Color green: rows A11:A20 where column A is "Li" and dates D11:D20 are in the next month.
 */

let formula1 =
  '=AND(A20="Zhang",YEAR(D20)=YEAR(TODAY()),MONTH(D20)=MONTH(TODAY()))'
let targetFormulaRange = Range('D20:D100')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
formatCondition1.SetFirstPriority()
formatCondition1.Interior.Pattern = xlPatternSolid
formatCondition1.Interior.Color = RGB(255, 0, 0)

let formula2 =
  '=AND(A11="Li",YEAR(D11)=YEAR(EOMONTH(TODAY(), 1)),MONTH(D11)=MONTH(EOMONTH(TODAY(), 1)))'
const formatCondition2 = Range('D11:D20').FormatConditions.Add(
  xlExpression,
  -1,
  formula2,
  ''
)
formatCondition2.SetLastPriority()
formatCondition2.Interior.Pattern = xlPatternSolid
formatCondition2.Interior.Color = RGB(0, 255, 0)
```

## SetLastPriority()

Sets evaluation order so this rule runs after all others.

The resulting priority equals the total number of rules on the sheet. Other rules with higher priority numbers are incremented by 1.

### Return value

undefined

### Example

```js
/**
 * Color red: rows A20:A100 where column A is "Zhang" and dates D20:D100 are in the current month.
 * Color green: rows A11:A20 where column A is "Li" and dates D11:D20 are in the next month.
 */

let formula1 =
  '=AND(A20="Zhang",YEAR(D20)=YEAR(TODAY()),MONTH(D20)=MONTH(TODAY()))'
let targetFormulaRange = Range('D20:D100')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
formatCondition1.SetFirstPriority()
formatCondition1.Interior.Pattern = xlPatternSolid
formatCondition1.Interior.Color = RGB(255, 0, 0)

let formula2 =
  '=AND(A11="Li",YEAR(D11)=YEAR(EOMONTH(TODAY(), 1)),MONTH(D11)=MONTH(EOMONTH(TODAY(), 1)))'
const formatCondition2 = Range('D11:D20').FormatConditions.Add(
  xlExpression,
  -1,
  formula2,
  ''
)
formatCondition2.SetLastPriority()
formatCondition2.Interior.Pattern = xlPatternSolid
formatCondition2.Interior.Color = RGB(0, 255, 0)
```
