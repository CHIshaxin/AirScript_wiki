---
title: "Range"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Range.html
updated: 1789700916000
hash: Bj-x7i_5
---

# Range

1 つのセル、1 行、1 列、または連続して選択されたセル ブロックを表します。

Range は、ワークシート上のセルの部分集合であり、まとめて読み取ったり変更したりできます。

以下に簡単な例を示します。Range オブジェクトのプロパティとメソッドは、次の一覧を参照してください。

```js
// セル A1:A10 を選択
let range = Application.Range('A1:A10')
// A1:A10 に行番号を設定
for (let i = 1; i <= 10; i++) {
  range.Item(i).Value = i.toString()
}
// 出力
printText(range) // 1...10
// 各セルに接頭辞と接尾辞を付与
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

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [Count](#count) | Number | 範囲内のセル数 |
| [Text](#text) | String | [読み取り専用] セルに表示される書式付きテキスト |
| [Value/Value2](#value) | any/[][]any | セルの値を取得または設定 |
| [FormatConditions](#formatConditions) | [FormatConditions](/docs/ja-JP/api/excel/workbook/FormatConditions.html) | 範囲の条件付き書式 |
| [Formula](#formula) | String | A1 形式の数式（暗黙の交差） |
| [FormulaArray](#formulaarray) | String | 範囲の配列数式を取得または設定 |
| [NumberFormat](#numberformat) | String | 表示形式の取得または設定 |
| [Hidden](#hidden) | Boolean | 行または列の非表示／表示 |
| [Interior.Color](#interior-color) | String | 16 進 RGB による塗りつぶしの色 |
| [HorizontalAlignment](#horizontalalignment) | [Enum.XlHAlign](/docs/ja-JP/api/excel/workbook/Enum.html#xlhalign) | 水平方向の配置 |
| [VerticalAlignment](#verticalalignment) | [Enum.XlVAlign](/docs/ja-JP/api/excel/workbook/Enum.html#xlvalign) | 垂直方向の配置 |
| [WrapText](#wraptext) | Boolean | 範囲内で折り返して全体を表示 |
| [IndentLevel](#indentlevel) | Number | セルのインデント |
| [MergeArea](#mergearea) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | セルを含む結合範囲 |
| [MergeCells](#mergecells) | Boolean | 範囲に結合セルが含まれるかどうか |
| [Cells](#cells) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 範囲内のセル |
| [Rows](#rows) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 範囲内の行 |
| [Columns](#columns) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 範囲内の列 |
| [EntireRow](#entirerow) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 範囲と交差する行全体 |
| [EntireColumn](#entirecolumn) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 範囲と交差する列全体 |
| [Row](#row) | Number | 範囲の先頭行番号 |
| [RowEnd](#rowend) | Number | 範囲の末尾行番号 |
| [Column](#column) | Number | 範囲の先頭列インデックス |
| [ColumnEnd](#columnend) | Number | 範囲の末尾列インデックス |
| [Borders](#borders) | [Borders](#borders) | 罫線のコレクション |

#### メソッド

| メソッド | 戻り値の型 | 説明 |
|---|---|---|
| [BorderAround()](#borderaround) | undefined | 罫線を追加し、新しい罫線に Color、LineStyle、Weight を設定 |
| [Each()](#each) | undefined | 範囲内の各セルを反復処理 |
| [Item()](#item) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 範囲内の指定インデックス位置のセル |
| [Offset()](#offset) | undefined | 行と列で範囲をオフセット |
| [Replace()](#replace) | undefined | セル内の文字列を置換 |
| [Delete()](#delete) | undefined | セル、行、または列を削除 |
| [Insert()](#insert) | undefined | セル、行、または列を挿入 |
| [InsertImage()](#insertimage) | undefined | セルに画像を挿入 |
| [Merge()](#merge) | undefined | セルを結合 |
| [UnMerge()](#unmerge) | undefined | セルの結合を解除 |
| [Address()](#address) | String | マクロ参照形式のアドレス文字列 |
| [AddComment()](#addcomment) | undefined | コメントを追加 |
| [ClearComments()](#clearcomments) | undefined | 範囲内のコメントをクリア |
| [Clear()](#clear) | undefined | 範囲の値と書式をクリア |
| [ClearContents()](#clearcontents) | undefined | 値のみクリア |
| [ClearFormats()](#clearformats) | undefined | 書式のみクリア |
| [ClearHyperlinks()](#clearhyperlinks) | undefined | ハイパーリンクの書式をクリア |
| [Contain()](#contain) | Boolean | 2 つの範囲が重なるか |
| [Copy()](#copy) | Boolean | 範囲をクリップボードにコピー |
| [Cut()](#cut) | Boolean | 範囲を切り取り先へ |
| [PasteSpecial()](#pastespecial) | undefined | オプション付きでクリップボードから貼り付け |
| [FillLeft()](#fillleft) | undefined | 右端の列から左方向に塗りつぶし |
| [FillRight()](#fillright) | undefined | 左端の列から右方向に塗りつぶし |
| [FillDown()](#filldown) | undefined | 上端の行から下方向に塗りつぶし |
| [FillUp()](#fillup) | undefined | 下端の行から上方向に塗りつぶし |
| [AutoFill()](#autofill) | undefined | 貼り付け先の範囲をオートフィル |
| [AutoFilter()](#autofilter) | undefined | 範囲にオートフィルター条件を適用 |
| [AutoFit()](#autofit) | undefined | 列幅または行の高さを自動調整 |
| [Select()](#select) | undefined | 範囲を選択 |
| [TextToColumns()](#texttocolumns) | undefined | 1 列のテキストを複数列に分割 |

## Count

範囲内の列数、行数、またはセル数。既定はセル数です。

#### データ型

Number - 列、行、またはセルの数

#### 例

```js
// セル範囲では Range.Count は Range.Cells.Count と等しい
let range = Application.Range('A1:B2')
console.log(range.Count) // 4
console.log(range.Cells.Count) // 4
console.log(range.Rows.Count) // 2
console.log(range.Columns.Count) // 2
```

## Text

[読み取り専用] 書式設定済みのセル テキスト

#### データ型

String - 書式設定済みのセル テキスト

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')

// [読み取り専用] 書式設定済みのセル テキスト
console.log(range.Text)

// Text は読み取り専用。セル内容を変更するときは Value を使用
range.Value = 'NewVaule'
console.log(range.Text) // NewValue
```

## Value/Value2

セルの値を取得または設定します。範囲の読み取りに対応します。

#### データ型

単一セル: any

範囲: [][]any

#### 例

```js
// Value プロパティ
let range = Application.Range('A1')

range.Value = 'A1'
// 出力 A1 value
console.log(range.Value) // A1

range = Range('A1:B2')

// A1:B2 に代入
range.Value = 'WebOffice'

// 出力 A1:B2 values
console.log(range.Value) // [["WebOffice","WebOffice"],["WebOffice","WebOffice"]]

// 一括代入
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]

// 出力 A1:B2 values
console.log(range.Value) // [["A1","B1"],["A2","B2"]]
```

## FormatConditions

FormatConditions コレクションは、Excel の条件付き書式を制御します。

#### データ型

[FormatConditions](/docs/ja-JP/api/excel/workbook/FormatConditions.html) - FormatConditions コレクション

#### 例

```js
// FormatConditions を取得
const formatConditions = Range('A:A').FormatConditions
// 書式条件の数
const count = formatConditions.Count

console.log(count)
```

## Formula

Formula は Value とは異なるルールでセル数式の読み書きを行います。

- 読み取り時: Formula は数式を返し、Text は表示値を返します。
- 書き込み時: Formula は 1 セルに適用されます。Value は範囲に対応します

#### データ型

String - 数式（暗黙の交差）

#### 例

```js
// A1=1、B1=2 に設定
Application.Range('A1').Value = '1'
Application.Range('B1').Value = '2'

// A2 の数式を =A1+B1 に設定
const range = Application.Range('A2')
range.Formula = '=A1+B1'
// A2 の表示値を読み取り: 3
console.log(range.Text) // 3
// A2 の数式を読み取り: =A1+B1
console.log(range.Formula) // =A1+B1
```

## FormulaArray

Formula と同様ですが、範囲にわたる配列の設定に対応します。

#### データ型

String - 範囲の配列数式

#### 例

```js
let range = Application.Range('A1:B2')

// A1:B2 の数式テキストを設定（例）
range.FormulaArray = '=Sum(A1:C3)'

range.Each(function (item) {
  console.log(item.Text) //=Sum(A1:C3), =Sum(A1:C3), =Sum(A1:C3), =Sum(A1:C3)
})
```

## NumberFormat

範囲の表示形式を取得または設定します。

取得時に範囲内で表示形式が異なる場合、先頭セルの書式が返されます。

WebOffice で表示形式を設定するときにコピーできる一般的な例:

- 標準: `G/General` (ロケール依存。ロケールによってはローカライズされた標準トークンを使用)
- 数値: `0.00_);[Red](0.00)`
- 通貨: `￥#,##0.00_);[Red](￥#,##0.00)`
- 会計: `_ ￥* #,##0.00_ ;_ ￥* -#,##0.00_ ;_ ￥* "-"??_ ;_ @_`
- 短い日付: `yyyy/m/d;@`
- 長い日付: ロケールに応じた長い日付パターン
- 時刻: `h:mm:ss;@`
- パーセント: `0.00%`
- 分数: `## ?/?`
- 指数: `0.00E+00`
- 文字列: `@`
- 桁区切りスタイル: `_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_`

上のいずれか 1 行を `Range.NumberFormat` にコピーします。例: 標準形式 `Range.NumberFormat = 'G/General'`。

#### データ型

String - 表示形式

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')

// 表示形式を取得
console.log('表示形式:', range.NumberFormat) // General（ロケール文字列の例）

// 書式を文字列に設定
range.NumberFormat = '@'
console.log('表示形式:', range.NumberFormat) // Text @
```

## Hidden

行または列の非表示／表示。書き込み専用（以前の非表示状態は読み取れません）。

#### データ型

Boolean - 行または列が非表示

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')

// 列全体
let entireColumn = range.EntireColumn

// 列を非表示
entireColumn.Hidden = true
```

## Interior.Color

塗りつぶしの色

#### データ型

Number - RGB による塗りつぶしの色

#### 例

```js
// Interior オブジェクト
let interior = Application.Range('A1').Interior

// 塗りつぶしの色を取得
console.log(interior.Color) // 16777215

// 塗りつぶしの色を赤に設定
interior.Color = RGB(255, 0, 0)
```

## HorizontalAlignment

水平方向の配置。次の値を使用します: [Enum.XlHAlign](/docs/ja-JP/api/excel/workbook/Enum.html#xlhalign).

#### データ型

[Enum.XlHAlign](/docs/ja-JP/api/excel/workbook/Enum.html#xlhalign) - 水平方向の配置

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')

// 出力 alignment
console.log(range.HorizontalAlignment) // 1

// 配置: 中央
range.HorizontalAlignment = Application.Enum.XlHAlign.xlHAlignCenter

// 出力 alignment
console.log(range.HorizontalAlignment) // -4108 (center)
```

## VerticalAlignment

垂直方向の配置。次の値を使用します: [Enum.XlVAlign](/docs/ja-JP/api/excel/workbook/Enum.html#xlvalign).

#### データ型

[Enum.XlVAlign](/docs/ja-JP/api/excel/workbook/Enum.html#xlvalign) - 垂直方向の配置

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')

// 配置: 下
range.VerticalAlignment = Application.Enum.XlVAlign.xlVAlignBottom

// 配置を取得
console.log(range.VerticalAlignment) //-4107 (bottom)
```

## IndentLevel

セルのインデント

#### データ型

Number - インデント段階

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')

// セルに値を書き込み
console.log(range.IndentLevel) // 0
range.IndentLevel = 10
console.log(range.IndentLevel) // 10
```

## WrapText

範囲の折り返して全体を表示するかどうかを取得または設定します

#### データ型

Boolean - 折り返して全体を表示する

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')
// 折り返して全体を表示するか取得
console.log('折り返し:', range.WrapText) // true
range.WrapText = false
console.log('折り返し:', range.WrapText) // false
```

## MergeArea

セルを含む結合された範囲

#### データ型

Range - Range オブジェクトとしての結合範囲

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1:D2')
// セルを結合
range.Merge()
// 結合範囲に値を設定
range.MergeArea.Value = 'merge'
```

## MergeCells

範囲に結合セルが含まれるかどうか

#### データ型

Boolean - 結合セルを含む

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1:D2')
// セルを結合
range.Merge()
// 範囲に結合セルが含まれるかどうか
console.log(range.MergeCells) // true
```

## Cells

範囲内のすべてのセルを Range として扱います（Range のメンバーを使用）。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 範囲内のセルのコレクション

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1:B2')

// Cells コレクション
let cells = range.Cells

// 先頭のセル
let cell = cells.Item(1)
console.log(cell.Address()) // $A$1
```

## Rows

範囲内の行を Range として扱います。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 範囲内の行

#### 例

```js
// A1:B2 は 2 行
console.log(Application.Range('A1:B2').Rows.Count) // 2

// A1:D4 は 4 行
console.log(Application.Range('A1:D4').Rows.Count) // 4
```

## Columns

範囲内の列を Range として扱います。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 範囲内の列

#### 例

```js
// A1:B2 は 2 列
console.log(Application.Range('A1:B2').Columns.Count) // 2

// A1:D4 は 4 列
console.log(Application.Range('A1:D4').Columns.Count) // 4
```

## EntireRow

範囲と交差する行全体

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 範囲と交差する行全体

#### 例

```js
// 1 行全体は 16384 セル
console.log(Application.Range('A1').EntireRow.Count) // 16384

// 2 行
console.log(Application.Range('A1:B2').EntireRow.Count) // 32768
```

## EntireColumn

範囲と交差する列全体

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 範囲と交差する列全体

#### 例

```js
// 1 列全体は 1048576 セル
console.log(Application.Range('A1').EntireColumn.Count) // 1048576

// 2 列
console.log(Application.Range('A1:B2').EntireColumn.Count) // 2097152
```

## Row

範囲の先頭行番号

#### データ型

Number - 範囲の先頭行番号

#### 例

```js
let range = Application.Range('A1:D4')
// A1:D4 の先頭行は 1
console.log(range.Row) // 1
```

## RowEnd

範囲の末尾行番号

#### データ型

Number - 範囲の末尾行番号

#### 例

```js
let range = Application.Range('A1:D4')
// A1:D4 の末尾行は 4
console.log(range.RowEnd) // 4
```

## Column

範囲の先頭列番号

#### データ型

Number - 範囲の先頭列番号

#### 例

```js
let range = Application.Range('A1:D4')
// A1:D4 の左端の列は列 1
console.log(range.Column) // 1
```

## ColumnEnd

範囲の末尾列番号

#### データ型

Number - 範囲の末尾列番号

#### 例

```js
let range = Application.Range('A1:D4')
// A1:D4 の右端の列は列 4
console.log(range.ColumnEnd) // 4
```

## Borders

Borders コレクション

#### データ型

[Border](/docs/ja-JP/api/excel/workbook/Border.html) コレクション - Border オブジェクト

#### 例

```js
// Borders コレクション
let borders = Application.Range('A1').Borders

// A1 の斜め下方向の罫線の色を赤に
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FF0000'
```

### Borders.Item()

A single border in the Borders コレクション for a range or style

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Index | Enum |  | はい | 取得する罫線。次を参照: [Enum.XlBordersIndex](/docs/ja-JP/api/excel/workbook/Enum.html#xlbordersindex) |

#### 戻り値の型

[Border](#border) - A single border in the Borders コレクション for a range or style

#### 例

```js
// Borders コレクション
let borders = Application.Range('A1').Borders

// A1 の斜め下方向の罫線の色を黄に
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FFFF00'
```

## BorderAround()

罫線を追加し、Color、LineStyle、Weight を設定します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| LineStyle | [LineStyle](/docs/ja-JP/api/excel/workbook/Enum.html#xllinestyle) | null | いいえ | 線種の XlLineStyle 定数 |
| BorderWeight | [BorderWeight](/docs/ja-JP/api/excel/workbook/Enum.html#xlborderweight) | null | いいえ | 罫線の太さ |
| Color | string | null | いいえ | 罫線の色（RGB）。例: `#ff0000` |

#### 例

```js
// B2 に赤の破線罫線
Range('B2').BorderAround(xlDash, xlHairline, '#ff0000')
// C3 に緑の一点鎖線罫線
Range('C3').BorderAround(xlDashDot, xlHairline, '#00ff00')
```

## Each()

各セルを反復処理します。結果を安定させるため、反復中に行／列の挿入・削除は避けてください。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| callback | Function | null | はい | Array.prototype.forEach に似たコールバック |

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1:D2')

// 各セルを走査
range.Each(function (item) {
  // セルアドレスを出力
  console.log(item.Address()) // $A$1...$D$2
})
```

## Item()

範囲内の指定インデックスに対応する Range を返します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| RowIndex | Number |  | はい | ColumnIndex を指定した場合は範囲内の相対行。それ以外は部分範囲へのインデックス |
| ColumnIndex | Number |  | いいえ | 範囲内の相対列 |

#### 戻り値の型

Range - 指定位置のセル

#### 例

```js
// Range オブジェクト
const range = Application.Range('A1:D2')
// 部分項目の例: B2
console.log(range.Item(2, 2).Address()) //$B$2
```

## Offset()

範囲をオフセットし、新しい Range を返します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| RowOffset | Number | 0 | いいえ | 行オフセット。正で下、負で上 |
| ColumnOffset | Number | 0 | いいえ | 列オフセット。正で右、負で左 |

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1:D2')
// 部分項目の例: B2
console.log(range.Address()) //$A$1:$D$2
// 範囲をオフセット
let newRange = range.Offset(2, 2)
// 部分項目の例: B2
console.log(newRange.Address()) //$C$3:$F$4
```

## Replace()

セル内の文字列を置換します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| What | String | （なし） | はい | 検索する文字列 |
| Replacement | String | （なし） | はい | 置換後の文字列 |
| LookAt | Enum | xlWhole | いいえ | XlLookAt: xlWhole または xlPart。例を参照 |
| SearchOrder | Enum | xlByRows | いいえ | XlSearchOrder: xlByRows または xlByColumns。例を参照 |
| MatchCase | Boolean | false | いいえ | true のとき大文字／小文字を区別 |
| MatchByte | Boolean | false | いいえ | true のとき全角と半角を区別 |

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1:D2')
// XiaoShan をすべて DaShan に置換（部分一致）
range.Replace('foo', 'bar', xlPart)
// セル値が完全一致するときのみ置換
range.Replace('foo', 'bar', xlWhole)
// 大文字／小文字を区別
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, true)
// 大文字／小文字を区別しない
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, false)
// 全角／半角を区別する
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, true, true)
// 全角／半角を区別しない
range.Replace('JinXiaoShan', 'JinShan', xlPart, xlByRows, true, false)
```

## Delete()

セル、行、または列を削除します。セルを削除すると残りが左にシフトします。

#### 例: 行の削除

```js
// Range オブジェクト
let range = Application.Range('B5:D10')
// この範囲がまたがる行 5～10 を削除
range.EntireRow.Delete()
```

#### 例: 列の削除

```js
let range = Application.Range('B5:D10')
// この範囲がまたがる列 2～4 を削除
range.EntireColumn.Delete()
```

## Insert()

セル、行、または列を挿入します。

#### 例: 行の挿入

```js
let range = Application.Range('B5:D10')
// 行 5 の上に 1 行挿入。元の行 5 は下にシフト
range.EntireRow.Insert()
```

#### 例: 列の挿入

```js
let range = Application.Range('B5:D10')
// 列 2 の左に 1 列挿入。元の列 2 は右にシフト
range.EntireColumn.Insert()
```

## InsertImage()

セルに画像を挿入します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| dataURL | string | undefined | はい | Base64 の data URL 文字列としての画像 |

#### 例

```js
// セル E1
const range = Range('E1')
// 対象セルに画像を挿入
range.InsertImage(
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII='
)
```

## Merge()

セルを結合します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Across | boolean | false | はい | true のとき、範囲内の各行を列方向に 1 セルに結合します。 |

#### 例

```js
const range = Application.Range('A1:D2')
// セルを結合
range.Merge()
```

## UnMerge()

セルの結合を解除します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| CancelCenter | Boolean | false | いいえ | 結合を解除するときに中央揃えもクリアするか |

#### 例

```js
const range = Application.Range('A1:D2')

// 結合を解除
range.UnMerge()
```

## Address()

マクロ参照形式のアドレス文字列。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| RowAbsolute | Boolean |  | いいえ | true のとき行部分を絶対参照で返す。既定は true |
| ColumnAbsolute | Boolean |  | いいえ | true のとき列部分を絶対参照で返す。既定は true |
| ReferenceStyle | Enum |  | いいえ | 参照形式。既定は xlA1。次を参照: [Enum.XlReferenceStyle](/docs/ja-JP/api/excel/workbook/Enum.html#xlreferencestyle) |
| External | Boolean |  | いいえ | true のとき外部参照、false のときローカル参照。既定は false |
| RelativeTo | Range |  | いいえ | 行と列がともに相対でスタイルが xlR1C1 のとき、相対アドレスの基準となる範囲 |

#### 戻り値の型

String - マクロ参照形式のアドレス文字列

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')

// アドレス文字列
console.log('address1:', range.Address())

console.log('address2:', range.Address(false, false))

console.log(
  'address3:',
  range.Address(true, true, Application.Enum.XlReferenceStyle.xlR1C1)
)
```

## AddComment()

コメントを追加します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Text | String |  | いいえ | コメントのテキスト |

#### 例

```js
let range = Application.Range('A1')
// A1 にコメントを追加
range.AddComment('WebOffice')
```

## ClearComments()

範囲内のコメントをクリアします。

#### 例

```js
let range = Application.Range('A1')
// 範囲内のコメントをクリアします。
range.ClearComments()
```

## Clear()

値と書式をクリアします。

#### 例

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

内容をクリアします。

#### 例

```js
let range = Application.Range('A1')
// 内容をクリアします。
range.ClearContents()
```

## ClearFormats()

書式をクリアします。

#### 例

```js
let range = Application.Range('A1')
// 書式をクリアします。
range.ClearFormats()
```

## ClearHyperlinks()

ハイパーリンクの書式をクリアします。

#### 例

```js
let range = Application.Range('A1')
// ハイパーリンクの書式をクリアします。
range.ClearHyperlinks()
```

## Contain()

2 つの範囲が重なるかどうか。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Range | Range |  | はい | 別の Range オブジェクト |

#### 戻り値の型

Boolean - 範囲が重なるかどうか

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1:D2')

// 2 つ目の Range オブジェクト
let newRange = Application.Range('A1:B4')

// 重なりをテスト
console.log(range.Contain(newRange))
```

## Copy()

範囲をクリップボードにコピーします。

#### 戻り値の型

Boolean - コピーが成功したか

#### 例

```js
// A1 をコピー
Range('A1').Copy()
// B1 に貼り付け
Range('B1').PasteSpecial()
```

## Cut()

範囲を切り取り、貼り付け先へ移動します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Destination | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 貼り付け先の範囲 |

#### 戻り値の型

Boolean - 切り取りが成功したか

#### 例

```js
// A1 を A2 へ移動
Range('A1').Cut(Range('A2'))
```

## PasteSpecial()

値、書式、数式などのオプションを指定してクリップボードの内容を貼り付けます。

注意

PasteSpecial() の前に範囲をクリップボードにコピーしてください。クリップボードが空のまま貼り付けるとエラーになります。

先に [Copy()](#copy) を実行し、続けて希望のオプションで PasteSpecial() を呼び出します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Paste | [XlPasteType](/docs/ja-JP/api/excel/workbook/Enum.html#xlpastetype) | xlPasteAll | いいえ | 貼り付けの種類。例: xlPasteAll または xlPasteValues |
| Operation | [XlPasteSpecialOperation](/docs/ja-JP/api/excel/workbook/Enum.html#xlpastespecialoperation) | xlPasteSpecialOperationNone | いいえ | 貼り付け演算。例: xlPasteSpecialOperationAdd |
| SkipBlanks | Boolean | false | いいえ | true のとき貼り付け時に元の空白セルをスキップ |
| Transpose | Boolean | false | いいえ | true のとき貼り付け時に行と列を入れ替えます |

#### 戻り値の型

undefined

#### 例 1

値のみ貼り付け。例: B1 を A1 に値として貼り付け。

```js
Range('B1').Copy()
Range('A1').PasteSpecial(xlPasteValues)
```

#### 例 2

A1:B2 から C1:D2 に書式を貼り付け。

```js
Range('A1:B2').Copy()
Range('C1:D2').PasteSpecial(xlPasteFormats)
```

#### 例 3

A1:B2 から A3:B4 に値と書式を貼り付け。

```js
Range('A1:B2').Copy()
Range('A3:B4').PasteSpecial(xlPasteAll)
```

#### 例 4

A1:B2 を C1:D2 に行列転置で貼り付け。

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

右端の列の値で左方向に塗りつぶします。

#### 例

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

左端の列の値で右方向に塗りつぶします。

#### 例

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

上端の行の値で下方向に塗りつぶします。

#### 例

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

下端の行の値で上方向に塗りつぶします。

#### 例

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

貼り付け先の範囲をオートフィルします。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Destination | Range |  | いいえ | 貼り付け先の範囲。元の範囲を含める必要があります |
| Type | Enum | Enum.XlAutoFillType.xlFillDefault | いいえ | 塗りつぶしの種類。次を参照: [Enum.XlAutoFillType](/docs/ja-JP/api/excel/workbook/Enum.html#xlautofilltype) |

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')
range.Value = 1
// 塗りつぶすセル
let fillRange = Application.Range('A1:A20')
// 貼り付け先の範囲をオートフィルします。
range.AutoFill(fillRange)
fillRange.Each(item => {
  console.log(item.Text()) // AutoFilled: 1,2,3,4...20
})
```

## AutoFilter()

範囲にオートフィルターを適用します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Field | Variant | null | いいえ | リスト左からのフィールド インデックス。左端のフィールドは 1 |
| Criteria1 | Variant | null | いいえ | 第 1 の条件。空白は "="、非空白は "<>" 。省略時はすべての値。xlTop10Items のときは件数を指定 |
| Operator | [XlAutoFilterOperator](/docs/ja-JP/api/excel/workbook/Enum.html#xlautofilteroperator) | null | いいえ | フィルター演算子。次を参照: [XlAutoFilterOperator](/docs/ja-JP/api/excel/workbook/Enum.html#xlautofilteroperator) |
| Criteria2 | Variant | null | いいえ | Criteria1 と Operator と組み合わせる第 2 の条件。日付フィルターでは Array(Level, Date) としても使用。Level は 0～2（年・月・日）、Date は範囲内の有効な日付 |
| SubField | Variant | null | いいえ | 構造化された型のフィールド（例: Population または Volume）。表示値のみのときは省略 |
| VisibleDropDown | Variant | true | いいえ | `true` のときフィールドのオートフィルター ドロップダウンを表示、`false` のとき非表示。既定は `true` |

#### 例 1

列 B の値で下から 11 行を表示。

```js
// 範囲を渡さないとき、フィルター領域には UsedRange が使われます
const filterRange = ActiveSheet.UsedRange
// 列 B
let filterField = 2
// フィルター対象のデータ列
const filterColumnRange = filterRange.Columns(filterField)
// 行数から見出し行を除く
const filterRangeCount = filterColumnRange.Rows.Count - 1
const bottom11Value = 11
const bottomNum =
  filterRangeCount < bottom11Value ? filterRangeCount : bottom11Value
let bottom11 = WorksheetFunction.Small(filterColumnRange, bottomNum)
filterRange.AutoFilter(filterField, '<=' + bottom11)
```

#### 例 2

列 1 が「1 Class」のとき、列 5 で値が 60 未満なら赤、90 超なら緑で塗りつぶし

```js
// 使用済み範囲をフィルター領域に使用
const filterRange = ActiveSheet.UsedRange
// 列 1 をフィルター
const filterColumn1 = 1
// 列 1 のフィルター値
let filterValue = '1 Class'
// フィルターを適用
filterRange.AutoFilter(filterColumn1, filterValue)
// 列 5 をフィルター
const filterColumn2 = 5
// 条件: 60 未満
filterValue = '<60'
// フィルターを適用
filterRange.AutoFilter(filterColumn2, filterValue)
// 見出し行を除いたデータ
const dataRange = ActiveSheet.UsedRange.Offset(1).Resize(
  ActiveSheet.UsedRange.Rows.Count - 1
)
// 列範囲
const columnRange = dataRange.Columns(filterColumn2)
// フィルター後の表示セル
const visibleRange = columnRange.SpecialCells(xlCellTypeVisible)
// 赤色を定義
const redColor = RGB(255, 0, 0)
// 赤に設定
visibleRange.Interior.Color = redColor
// 列 5 のフィルターを解除
filterRange.AutoFilter(filterColumn2)
// 条件: 90 超
const filterValue2 = '>90'
// フィルターを適用
filterRange.AutoFilter(filterColumn2, filterValue2)
// フィルター後の表示セル
const visibleRange2 = columnRange.SpecialCells(xlCellTypeVisible)
// 緑色を定義
const greenColor = RGB(0, 255, 0)
// 緑に設定
visibleRange2.Interior.Color = greenColor
// 列 1 のフィルターを解除
filterRange.AutoFilter(filterColumn1)
// 列 5 のフィルターを解除
filterRange.AutoFilter(filterColumn2)
```

#### 例 3

列 A が指定日付に一致し列 C が 10000 超の行をフィルター

```js
// 既定のフィルター範囲は UsedRange
const filterRange = ActiveSheet.UsedRange
// 列 A をフィルター
let filterField = 1
// 日付条件の配列
let criteria = new Array()
// 日付レベルの列挙
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// 日条件 2023/4/10 を追加
criteria.push(FilterDateTimeEnum.DAY, '2023/4/10')
// 月条件 2022 年 12 月を追加
criteria.push(FilterDateTimeEnum.MONTH, '2022/12/1')
// フィルターを適用
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
// 列 C をフィルター
filterField = 3
// 条件: > 10000
let criteria2 = '>10000'
// フィルターを適用
filterRange.AutoFilter(filterField, criteria2)
```

#### 例 4

列 A を列挙した年／月／日の組み合わせでフィルター

```js
// 既定のフィルター範囲は UsedRange
const filterRange = ActiveSheet.UsedRange
// 列 A をフィルター
let filterField = 1
// 日付条件の配列
let criteria = new Array()
// 日付レベルの列挙
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// 日条件 2023/4/10 を追加
criteria.push(FilterDateTimeEnum.YEAR, '2023/1/1')
// 2022 年 3 月を追加
criteria.push(FilterDateTimeEnum.MONTH, '2022/3/1')
// 2022 年 4 月を追加
criteria.push(FilterDateTimeEnum.MONTH, '2022/4/1')
// 2022-05-01 を追加
criteria.push(FilterDateTimeEnum.DAY, '2022/5/1')
// 2022-05-02 を追加
criteria.push(FilterDateTimeEnum.DAY, '2022/5/2')
// フィルターを適用
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
```

#### 例 5

列 D が当四半期、列 C が年初来、列 A が Product A の行を表示

```js
// 既定のフィルター範囲は UsedRange
const filterRange = ActiveSheet.UsedRange
// 列 D をフィルター
let filterField = 4
// 四半期フィルター（今／前／次四半期）
let criteria = xlFilterThisQuarter
// フィルターを適用
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// 列 C をフィルター
filterField = 3
// 年初来フィルター
criteria = xlFilterYearToDate
// フィルターを適用
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// 列 A をフィルター
filterField = 1
// 値 Product A
criteria = 'ProductA'
// フィルターを適用
filterRange.AutoFilter(filterField, criteria)
```

#### 例 6

列 2 が Zhang を含むか Li で始まる行をフィルター

```js
// 既定のフィルター範囲は UsedRange
const filterRange = ActiveSheet.UsedRange
// 列 2 をフィルター
let filterField = 2
// Zhang を含む
let criteria = '=*Zhang*'
// Li で始まる
let criteria2 = '=Li*'
// OR で適用
filterRange.AutoFilter(filterField, criteria, xlOr, criteria2)
```

#### 例 7

列 A の値で上位 5 行を表示

```js
// 既定のフィルター範囲は UsedRange
const filterRange = ActiveSheet.UsedRange
// 列 A をフィルター
let filterField = 1
// データ列の範囲
const filterColumnRange = filterRange.Columns(filterField)
// 行数から見出し行を除く
const filterRangeCount = filterColumnRange.Rows.Count - 1
// 上位 5 件の値を取得
const top5Value = 5
// 利用可能行数に合わせてクランプ
const topNum = filterRangeCount < top5Value ? filterRangeCount : top5Value
// 5 番目に大きい値
let top5 = WorksheetFunction.Large(filterColumnRange, topNum)
// フィルターを適用
filterRange.AutoFilter(filterField, '>=' + top5)
```

#### 例 8

列 C が Zhang の表示行について列 F のフォントを楷体に

```js
// 使用済み範囲をフィルター領域に使用
const filterRange = ActiveSheet.UsedRange
// 列 3 をフィルター
const filterColumn = 3
// 値 Zhang
const filterValue = ['Zhang']
// フィルターを適用
filterRange.AutoFilter(filterColumn, filterValue)

// フィルター結果を取得
const filteredRange =
  ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible)

// 列 F の表示セル
const filteredRangeFColumn =
  ActiveSheet.AutoFilter.Range.Columns('F').SpecialCells(xlCellTypeVisible)

// 列 F に楷体を適用
filteredRangeFColumn.Font.Name = 'KaiTi'
```

#### 例 9

列 2 が Li かつ列 4 が先週の行を表示

```js
// 既定のフィルター範囲は UsedRange
const filterRange = ActiveSheet.UsedRange
// 列 2 をフィルター
let filterField = 2
// 値 Li
let criteria = 'Li'
// フィルターを適用
filterRange.AutoFilter(filterField, criteria)
// 列4をフィルター
filterField = 4
// 先週フィルター
criteria = xlFilterLastWeek
// フィルターを適用
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
```

## AutoFit()

列幅または行の高さを自動調整します。

[Columns](#column) または [Rows](#rows) に対して呼び出す必要があります。それ以外では効果がありません。

列幅の 1 単位は、標準スタイルで 1 文字の幅に相当します。

#### 例

```js
// 列幅を自動調整
const range = Application.Range('A1')
range.Columns.AutoFit()

// 行 1 の高さを自動調整
const range = Application.Range('A1')
range.Rows.AutoFit()
```

## Select()

範囲を選択します。Application.Selection を更新します。

#### 例

```js
// Range オブジェクト
let range = Application.Range('A1')
range.Select()
let selection = Application.Selection
console.log(selection.Address())
```

## TextToColumns()

区切り文字を使ってセルのテキストを複数列に分割します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Destination | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | いいえ | 分割先の列。省略時は元のセルを上書き |
| DataType | [XlTextParsingType](/docs/ja-JP/api/excel/workbook/Enum.html#xltextparsingtype) | xlDelimited | いいえ | 解析するテキストの形式 |
| TextQualifier | [XlTextQualifier](/docs/ja-JP/api/excel/workbook/Enum.html#xltextqualifier) | xlTextQualifierDoubleQuote | いいえ | テキストの区切り記号 |
| ConsecutiveDelimiter | Boolean | false | いいえ | 連続する区切りを 1 つとして扱う |
| Tab | Boolean | false | いいえ | タブを区切りとして使用 |
| Semicolon | Boolean | false | いいえ | セミコロンを区切りとして使用 |
| Comma | Boolean | false | いいえ | コンマを区切りとして使用 |
| Space | Boolean | false | いいえ | スペースを区切りとして使用 |
| Other | Boolean | false | いいえ | `OtherChar` を区切りとして使用 |
| OtherChar | String | `''` | いいえ | カスタム区切り文字。既定は空 |
| FieldInfo | Array | null | いいえ | 列ごとの解析オプションの配列 |
| DecimalSeparator | String | システム設定 | いいえ | 数値解析時の小数点 |
| ThousandsSeparator | String | システム設定 | いいえ | 数値解析時の桁区切り |
| TrailingMinusNumbers | String | null | いいえ | マイナス記号で始まる数値の扱い |

ヒント

`FieldInfo` について: この配列は `DataType` に応じて各列をどう解析するかを記述します。

区切りデータでは、各要素は 2 要素の配列です: 列インデックス（1 始まり）と [XlColumnDataType](/docs/ja-JP/api/excel/workbook/Enum.html#xlcolumndatatype) の値。エントリのない列は標準。固定幅データでは第 1 要素が開始文字インデックスです。

#### 戻り値の型

undefined

#### 例 1

A1:A3 をタブ、列 B をセミコロン、C1 をコンマ、D1 をスペースで分割。

```js
// 1. タブ区切り A1:A3
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
// 2. セミコロン区切りの列 B
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
// 3. コンマ区切りの C1
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
// 4. スペース区切りの D1
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

#### 例 2

OtherChar を使って A1 を `@`、A2 を `#`、A3 を `*` で分割。

```js
// 1. A1 で OtherChar '@'
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
// 2. A2 で OtherChar '#'
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
// 3. A3 で OtherChar '*'
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

#### 例 3

A1:A10 をスペースとセミコロンで B1 へ。列 B をコンマ、タブ、`=` で列 D へ。列 C をタブ、コンマ、セミコロン、スペースで分割。

```js
// 1. スペース＋セミコロンを B1 へ
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
// 2. 列 B をコンマ・タブ・'=' で列 D へ
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
// 3. 列 C をタブ・コンマ・セミコロン・スペースで
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

#### 例 4

A1 をコンマ分割（連続区切りを 1 つとして）。B1:B5 を `&` とタブで分割（連続区切りを結合）。列 C を `!` で分割（連続区切りを別々に）。

```js
// 1. コンマ、ConsecutiveDelimiter true
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
// 2. OtherChar '&' とタブ、ConsecutiveDelimiter true
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
// 3. OtherChar '!'、ConsecutiveDelimiter false
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

#### 例 5

A1 をスペースで B1 に分割し FieldInfo で標準・文字列・YMD 日付。C1 をセミコロンで C3 に分割し DMY、MYD、DYM、YDM の日付列に。

```js
// 1. スペース区切り、FieldInfo [列, xlColumnDataType]
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
// 2. セミコロン区切り、列ごとの日付書式
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
