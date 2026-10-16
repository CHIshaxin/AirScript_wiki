---
title: "FormatCondition"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/FormatCondition.html
updated: 1789700916000
hash: WMTvssof
---

# FormatCondition

範囲内の 1 つの条件付き書式です。

FormatCondition オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [AppliesTo](#appliesto) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 書式ルールが適用される Range |
| [Borders](#borders) | [Border](/docs/ja-JP/api/excel/workbook/Border.html) | Borders コレクション |
| [Font](#font) | [Font](/docs/ja-JP/api/excel/workbook/Font.html) | Font オブジェクト |
| [Formula1](#formula1) | 文字列 | 条件または入力規則の値・式 |
| [Formula2](#formula2) | 文字列 | 入力規則や 2 値の間の条件の第 2 要素 |
| [Interior](#interior) | Interior | オブジェクトの塗りつぶし |
| [NumberFormat](#numberformat) | 文字列 | ルールが真のときの表示形式 |
| [Operator](#operator) | [XlFormatConditionOperator](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditionoperator) | 条件付き書式の演算子 |
| [Priority](#priority) | 数値 | このルールの優先度 |
| [Type](#type) | [XlFormatConditionType](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditiontype) | 条件オブジェクトの種類 |

#### メソッド

| メソッド | 戻り値 | 説明 |
|---|---|---|
| [Modify()](#modify) | undefined | 既存のルールを変更する |
| [ModifyAppliesToRange()](#modifyappliestorange) | undefined | このルールが適用される範囲を設定する |
| [SetFirstPriority()](#setfirstpriority) | undefined | このルールの優先度を 1 に設定する（最初に評価） |
| [SetLastPriority()](#setlastpriority) | undefined | このルールの優先度を最後に設定する（他を繰り上げ） |

## AppliesTo

書式ルールが適用される Range を返します。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - Range オブジェクト

#### 例

```js
// A が 2 のときの B の平均。A が 2 のときに B がその平均より大きいセルを強調
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

範囲またはスタイル（条件付き書式の罫線を含む）の Borders を返します。

#### データ型

[Border](/docs/ja-JP/api/excel/workbook/Border.html) コレクション

#### 例

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

書式設定された範囲の Font オブジェクトを返します。

#### データ型

[Font](/docs/ja-JP/api/excel/workbook/Font.html) - Font オブジェクト

#### 例

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

条件または入力規則に関連付けられる値または式です。

Formula1 は定数、文字列、参照、または式のいずれかになり得ます。

#### データ型

文字列 - 値または式

#### 例

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

条件または入力規則の第 2 部分です。

Operator が xlBetween または xlNotBetween のときに使用します。定数、文字列、参照、または式のいずれかになり得ます。

#### データ型

文字列 - 値または式

#### 例

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

Interior オブジェクト（塗りつぶし）を返します。

#### データ型

Interior - Interior オブジェクト

#### 例

```js
const range = Range('B2:G18')
range.FormatConditions.Delete()
let formatCondition = range.FormatConditions.AddUniqueValues()
formatCondition.DupeUnique = xlDuplicate
const redColor = RGB(255, 0, 0)
formatCondition.Interior.Color = redColor
```

## NumberFormat

ルールの評価結果が True のときに適用される表示形式を取得または設定します。

「セルの書式設定」の「表示形式」タブと同じ形式記号を使用します。

「標準」などの組み込み形式や、カスタムコードを使用できます。

#### データ型

文字列 - 表示形式

#### 例

```js
let targetFormulaRange = Range('B:B')
let formatCondition1 = targetFormulaRange.FormatConditions
formatCondition1.NumberFormat = '0.0'
```

## Operator

条件付き書式の演算子を返します。

#### データ型

[XlFormatConditionOperator](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditionoperator) - 演算子の列挙型

#### 例

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
console.log(betweenCondition.Operator)
```

## Priority

ルールの優先度を取得または設定します。複数のルールがある場合、優先度が評価順を制御します。

#### データ型

数値 - 優先度

#### 例

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
console.log(betweenCondition.Priority)
```

## Type

オブジェクトの種類を返します。

[XlFormatConditionType](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditiontype) のいずれか:

- xlCellValue
- xlExpression

#### データ型

[XlFormatConditionType](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditiontype) - 条件の種類

#### 例

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
console.log(betweenCondition.Type)
```

## Modify()

既存の条件付き書式を変更します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Type | [XlFormatConditionType](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditiontype) | null | はい | セル値ベースまたは式ベース |
| Operator | [XlFormatConditionOperator](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditionoperator) | null | いいえ | Type が xlExpression のときは無視される |
| Formula1 | 文字列 | null | いいえ | 値または式 |
| Formula2 | 文字列 | null | いいえ | 第 2 の値または式 |

### 戻り値

undefined

#### 例

```js
const formatconditions = Range('F:F').FormatConditions
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
betweenCondition.Modify(xlCellValue, xlBetween, '=90', '=110')
```

## ModifyAppliesToRange()

このルールが適用されるセル範囲を設定します。

A1 参照スタイルであり、FormatConditions コレクションを所有するワークシート上に完全に収まっている必要があります。範囲演算子（コロン）、交差（スペース）、和集合（カンマ）が使用可能です。通貨記号は無視されます。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | このルールが適用される範囲 |

### 戻り値

undefined

### 例

```js
const formatConditions = Range('B:B').FormatConditions
const destRange = Range('A:A')
const count = formatConditions.Count
for (let i = 1; i <= count; i++) {
  formatConditions.Item(i).ModifyAppliesToRange(destRange)
}
```

## SetFirstPriority()

このルールの優先度を 1 に設定し、他のすべてより先に評価されるようにします。

すでに先頭でない場合、他のルールの優先度は 1 ずつ繰り上がります。

注意

条件付き書式の優先度は、ワークシート単位で評価されます。

### 戻り値

undefined

### 例

```js
/**
 * 赤: A 列が "Zhang" かつ D20:D100 の日付が今月の行（A20:A100）。
 * 緑: A 列が "Li" かつ D11:D20 の日付が翌月の行（A11:A20）。
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

評価順を、このルールが他のすべての後に来るように設定します。

結果の優先度は、シート上のルールの合計数に等しくなります。より大きい優先度番号を持つ他のルールは 1 ずつ繰り上がります。

### 戻り値

undefined

### 例

```js
/**
 * 赤: A 列が "Zhang" かつ D20:D100 の日付が今月の行（A20:A100）。
 * 緑: A 列が "Li" かつ D11:D20 の日付が翌月の行（A11:A20）。
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
