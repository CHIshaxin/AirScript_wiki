---
title: "Sort"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Sort.html
updated: 1789700916000
hash: BGwVWCc1
---

# Sort

Sort オブジェクトです。対象の範囲を設定してから並べ替えを適用します。

#### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [Header](#header) | [XlYesNoGuess](Enum.md#xlyesnoguess) | 先頭行が見出し行かどうか |
| [MatchCase](#matchcase) | Boolean | 大文字／小文字を区別する並べ替え |
| [Orientation](#orientation) | [XlSortOrientation](Enum.md#xlsortorientation) | 並べ替えの向き |
| [Rng](#rng) | [Range](Range.md) | 並べ替え対象の値を含む範囲 |
| [SortFields](#sortfields) | [SortFields](https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/SortFields.html) | この **Sort** オブジェクトの並べ替えフィールドのコレクション |
| [SortMethod](#sortmethod) | [XlSortMethod](Enum.md#xlsortmethod) | 東アジア語の並べ替え方法（読み／画数） |

#### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [Apply()](#apply) | undefined | 現在の並べ替え状態で範囲を並べ替えます |
| [SetRange()](#setrange) | undefined | 並べ替え対象の範囲を設定します |

#### 使用例

ワークシート上でよくある並べ替えの例です。

例 1. 列 C を降順、列 D を昇順で並べ替え

```js
// 現在の使用済み範囲
const range = ActiveSheet.UsedRange
// Sort オブジェクト
const sort = ActiveSheet.Sort
// 並べ替えフィールド
const sortFields = sort.SortFields
// 既存のキーをクリア
sortFields.Clear()
// 列 C を値で降順
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
// 列 D を値で昇順
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
// 見出し行: xlGuess は自動、xlYes は見出しあり、xlNo は見出しなし
sort.Header = xlGuess
// 大文字／小文字の区別
sort.MatchCase = false
// 東アジアの並べ替え: xlPinYin または xlStroke
sort.SortMethod = xlPinYin
// xlSortColumns と xlSortRows
sort.Orientation = xlSortColumns
// Apply の前に SetRange が必要
sort.SetRange(range)
sort.Apply()
```

例 2. 列 G のセル色が赤い行を上に、列 F の赤を下に

```js
// 現在の使用済み範囲
const range = ActiveSheet.UsedRange
// Sort オブジェクト
const sort = ActiveSheet.Sort
// 並べ替えフィールド
const sortFields = sort.SortFields
// 既存のキーをクリア
sortFields.Clear()
// 列 G: 赤を上に
// 引数: キー範囲、並べ替えの基準（値／フォント色／セル色）、順序（昇順／降順）
const sortField1 = sortFields.Add(
  Range('G:G').Item(1, 1),
  xlSortOnCellColor,
  xlAscending
)
// 並べ替えの色は赤
sortField1.SortOnValue.Color = RGB(255, 0, 0)
// 列 F: 赤を下に
const sortField2 = sortFields.Add(
  Range('F:F').Item(1, 1),
  xlSortOnCellColor,
  xlDescending
)
// 並べ替えの色は赤
sortField1.SortOnValue.Color = RGB(255, 0, 0)
sort.Header = xlGuess
sort.MatchCase = false
sort.SortMethod = xlPinYin
sort.Orientation = xlSortColumns
sort.SetRange(range)
sort.Apply()
```

例 3. シートに列 A～H があるとき、列 B の値が小さい順に並べ替え

```js
// Sort オブジェクト
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
// 値で昇順に並べ替え
sortFields.Add(sortKeyRange, xlSortOnValues, xlAscending)

// 並べ替えを適用する範囲
const applyBeginRow = ActiveSheet.UsedRange.Row
const applyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const applyRange = ActiveSheet.Range(`A${applyBeginRow}:H${applyEndRow}`)
// Apply の前に SetRange
sort.SetRange(applyRange)
sort.Apply()
```

例 4. シートに列 A～H があるとき、列 A を画数の降順で並べ替え

```js
// Sort オブジェクト
const sort = ActiveSheet.Sort

// キー範囲
const keyColumn = 'A'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow

// 値で降順
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
// 画数による並べ替え
sort.SortMethod = xlStroke

// 並べ替え範囲
const beginRow = ActiveSheet.UsedRange.Row
const endRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

sort.SetRange(ActiveSheet.Range(`A${beginRow}:H${endRow}`))
sort.Apply()
```

例 5. シートに列 A～H があるとき、列 B のセル色で並べ替え（塗りつぶしなしを上に）

```js
// キー範囲
const keyColumn = 'B'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow
let keyRange = ActiveSheet.Range(keyRangeStr)

// セル色で降順に並べ替え
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

先頭行に見出し行の情報が含まれるかどうか。読み書き可能。[XlYesNoGuess](Enum.md#xlyesnoguess)。

既定値は xlNo。xlGuess で Excel に見出し行の検出を任せます。

#### データ型

[XlYesNoGuess](Enum.md#xlyesnoguess) - 見出し行の扱い

## MatchCase

`true` のとき大文字／小文字を区別して並べ替え、`false` のとき区別しません。読み書き可能です。

#### データ型

Boolean - 大文字／小文字の区別

## Orientation

並べ替えの向きです。読み書き可能。[XlSortOrientation](Enum.md#xlsortorientation)。

#### データ型

[XlSortOrientation](Enum.md#xlsortorientation) - 並べ替えの向き

## Rng

並べ替え対象の値の範囲を返します。読み取り専用です。

#### データ型

[Range](Range.md) - 並べ替え範囲

## SortFields

この **Sort** オブジェクトの並べ替えフィールドのコレクションです。読み取り専用です。

### Count

コレクション内のオブジェクト数を返します。読み取り専用です。

#### データ型

Number - オブジェクト数

### Add()

新しい並べ替えフィールドを作成し、SortField オブジェクトを返します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Key | [Range](Range.md) | null | はい | キー範囲 |
| SortOn | [XlSortOn](Enum.md#xlsorton) | null | いいえ | 並べ替えの基準となるプロパティ |
| Order | [XlSortOrder](Enum.md#xlsortorder) | null | いいえ | 並べ替え順 |
| CustomOrder | Variant | null | いいえ | カスタム リストの使用 |
| DataOption | [XlSortDataOption](Enum.md#xlsortdataoption) | null | いいえ | データ オプション |

#### 戻り値

[SortField](SortField.md) - この **Sort** オブジェクトの並べ替えフィールド

### Clear()

すべての SortField オブジェクトをクリアします。新しいキーを追加する前に呼び出します。

#### 戻り値

undefined

### Item()

[SortField](SortField.md) オブジェクトを返します。読み取り専用です。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Index | Number / String | null | はい | 1 から始まるインデックス |

#### 戻り値

[SortField](SortField.md) - この **Sort** オブジェクトの並べ替えフィールド

## SortMethod

東アジア語の並べ替え方法（読みまたは画数）です。読み書き可能。[XlSortMethod](Enum.md#xlsortmethod)。

#### データ型

[XlSortMethod](Enum.md#xlsortmethod) - 東アジア語の並べ替え方法

## Apply()

現在の並べ替え状態で範囲を並べ替えます。

注意

並べ替えルールを適用する前に [SetRange](#setrange) を呼び出してください。呼び出さないと並べ替えは反映されません。

#### 戻り値

undefined

## SetRange()

並べ替え対象の範囲を設定します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Range | [Range](Range.md) | null | はい | この **Sort** オブジェクトが並べ替える範囲 |

#### 戻り値

undefined
