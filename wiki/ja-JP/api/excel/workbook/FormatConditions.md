---
title: "FormatConditions"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/FormatConditions.html
updated: 1789700916000
hash: DnAeILn6
---

# FormatConditions

FormatConditions コレクションは、Excel の条件付き書式を制御します。

条件付き書式は、セルの値、数式、その他のルールに基づいて書式を適用し、データを読み取りやすくします。

#### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [Count](#count) | 数値 | FormatConditions コレクション内のオブジェクト数 |

#### メソッド

| メソッド | 戻り値 | 説明 |
| --- | --- | --- |
| [Add()](#add) | [FormatCondition](FormatCondition.md) | コレクションに条件付き書式を追加する |
| [AddAboveAverage()](#addaboveaverage) | AboveAverage | 範囲の新しい AboveAverage ルール |
| [AddIconSetCondition()](#addiconsetcondition) | IconSetCondition | 範囲のアイコンセット条件付き書式 |
| [AddColorScale()](#addcolorscale) | ColorScale | 相対的なセル値をグラデーションで示すカラースケール ルール |
| [AddTop10()](#addtop10) | Top10 | 順位または割合で最大・最小の値を強調表示する |
| [AddUniqueValues()](#adduniquevalues) | UniqueValues | 範囲の新しい UniqueValues ルール |
| [Delete()](#delete) | undefined | 範囲の条件付き書式を削除する |
| [Item()](#item) | [FormatCondition](FormatCondition.md) | コレクションから 1 つの条件付き書式を返す |

## Count

FormatConditions コレクション内のオブジェクト数です。

#### データ型

数値 - オブジェクト数

#### 例

```js
const formatConditions = Range('A:A').FormatConditions
const count = formatConditions.Count

console.log(count)
```

## Add()

コレクションに条件付き書式を追加します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Type | [XlFormatConditionType](Enum.md#xlformatconditiontype) | null | はい | 条件の種類 |
| Operator | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | null | はい | 演算子。Type が `xlExpression` のときは無視される |
| Formula1 | 文字列 | null | はい | 値または式: 定数、文字列、参照、または式 |
| Formula2 | 文字列 | null | いいえ | Operator が `xlBetween` または `xlNotBetween` のときの第 2 要素（それ以外では無視）。定数、文字列、参照、または式 |

#### 戻り値

[FormatCondition](FormatCondition.md) - 追加された条件付き書式

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
formatCondition1.SetFirstPriority()
formatCondition1.Font.Strikethrough = true
```

## AddAboveAverage()

範囲の新しい AboveAverage オブジェクトを返します。

AboveAverage は、平均や標準偏差に対して上または下の値を強調します（人事評価など）。

#### 戻り値

AboveAverage オブジェクト

#### 例

```js
const formatconditions = Range('D:D').FormatConditions
const aboveAverageCondition = formatconditions.AddAboveAverage()
aboveAverageCondition.AboveBelow = xlAboveAverage
aboveAverageCondition.SetFirstPriority()
aboveAverageCondition.Font.Color = RGB(255, 255, 0)
```

## AddIconSetCondition()

範囲のアイコンセット ルール用の新しい IconSetCondition を返します。

アイコンセットはデータに注釈を付け、しきい値で 3〜5 段階に分割します。

#### 戻り値

IconSetCondition オブジェクト

#### 例

```js
// D 列と F 列にアイコンセット: >=130、90〜130、<=90

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

新しい ColorScale オブジェクトを返します。相対的なセル値をグラデーションで示します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| ColorScaleType | 数値 | null | はい | 2 は 2 色スケール、3 は 3 色スケール |

#### 戻り値

ColorScale オブジェクト

#### 例

```js
// B 列に 2 色スケール: 最小は赤、最大は緑

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

範囲内の上位または下位の値用の Top10 オブジェクトを返します。

Top10 は、順位または割合で最大・最小の値を強調するために使います。

#### 戻り値

Top10 オブジェクト

#### 例

```js
// L 列の下位 5 件を緑のフォントで強調

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

範囲の新しい UniqueValues オブジェクトを返します。

UniqueValues は、一意のセルまたは重複セルを強調するために使います。

#### 戻り値

UniqueValues オブジェクト

#### 例

```js
// B2:G18 の重複を赤の塗りつぶしで強調

const range = Range('B2:G18')
range.FormatConditions.Delete()
let formatCondition = range.FormatConditions.AddUniqueValues()
formatCondition.DupeUnique = xlDuplicate
const redColor = RGB(255, 0, 0)
formatCondition.Interior.Color = redColor
```

## Delete()

範囲の条件付き書式を削除します。

#### 戻り値

undefined

```js
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Delete()
```

## Item()

コレクションから 1 つの条件付き書式を返します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Index | 数値 / 文字列 | null | はい | コレクション内の 1 始まりのインデックス |

#### 戻り値

[FormatCondition](FormatCondition.md) オブジェクト

#### 例

```js
const formatConditions = Range('B:B').FormatConditions
const destRange = Range('A:A')
const count = formatConditions.Count
for (let i = 1; i <= count; i++) {
  formatConditions.Item(i).ModifyAppliesToRange(destRange)
}
```
