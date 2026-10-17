---
title: "Sheet"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Sheet.html
updated: 1789700916000
hash: nT3xMfd_
---

# Sheet

ブック（Workbook）内の単一のワークシート（Sheet）オブジェクトです。

Sheet オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

#### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [Id](#id) | String | このワークシートの ID |
| [Name](#name) | String | このワークシートの名前 |
| [Index](#index) | Number | すべてのシート間でのこのシートのインデックス |
| [Cells](#cells) | [Range](Range.md) | このシート上のすべてのセル |
| [Columns](#columns) | [Range](Range.md) | このシート上のすべての列 |
| [Rows](#rows) | [Range](Range.md) | このシート上のすべての行 |
| [UsedRange](#usedrange) | [Range](Range.md) | このシートの使用済み範囲 |
| [Visible](#visible) | Boolean | このシートを表示するかどうか |
| [Type](#type) | String | このワークシートの種類 |
| [Hyperlinks](#hyperlinks) | [Hyperlinks](#hyperlinks) | このシート上のすべてのハイパーリンク |
| [Shapes](#shapes) | [Shapes](#shapes) | このシート上のすべての Shape |
| [Sort](#sort) | [Sort](Sort.md) | このシート上の Sort オブジェクト |

#### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [Range()](#range) | [Range](Range.md) | セルまたはセル範囲 |
| [Cells()](#cells-1) | [Range](Range.md) | このシート上の 1 セル |
| [Activate()](#activate) | undefined | ワークシートをアクティブにする |
| [Move()](#move) | undefined | ワークシートを移動する |
| [Delete()](#delete) | undefined | ワークシートを削除する |

## Id

ワークシートの ID を取得します。

#### データ型

String - ワークシートの ID

#### 例

```js
const sheet = Application.ActiveSheet
// アクティブなワークシートの ID を出力
console.log(sheet.Id)
```

## Name

ワークシートの名前を取得または設定します。

#### データ型

String - すべてのシート間でのこのシートの名前

#### 例

```js
const sheet = Application.ActiveSheet
// アクティブなワークシートの名前を出力
console.log(sheet.Name) // Sheet2

// 現在のシート名を WPS WebOffice に変更
sheet.Name = 'WPS WebOffice'
```

## Index

シートのインデックス。すべてのワークシート間での位置です。

#### データ型

String - すべてのワークシート間でのインデックス

#### 例

```js
const sheet = Application.ActiveSheet
// アクティブなワークシートのインデックスを出力
console.log(sheet.Index) // 1
```

## Cells

シート上のすべてのセルを Range として扱うコレクションです（Range のプロパティとメソッドを使用します）。

#### データ型

[Range](Range.md) - シート上のすべてのセル

#### 例

```js
const sheet = Application.ActiveSheet
// アクティブなシートのセル全体のアドレス範囲を出力
console.log(sheet.Cells.Address()) // $A$1:$XFD$1048576
```

## Columns

シート上のすべての列を Range として扱います。[Range](Range.md) を参照してください。

#### データ型

[Range](Range.md) - シート上のすべての列

#### 例

```js
// このシートの列数を出力
const sheet = Application.ActiveSheet
console.log(sheet.Columns.Count) //16384
```

## Rows

シート上の行を Range として扱います（Range のプロパティとメソッドを使用します）。

#### データ型

[Range](Range.md) - シート上のすべての行

#### 例

```js
const sheet = Application.ActiveSheet
console.log(sheet.Rows.Count) // 1048576
```

## UsedRange

シートの使用済み範囲（実際に使用されている領域）を Range として扱います（Range のプロパティとメソッドを使用します）。

#### データ型

[Range](Range.md) - シート上のセルまたは範囲

#### 例

```js
const UsedRange = Application.ActiveSheet.UsedRange
// 使用済み範囲の境界を出力（ここでは 10×10 の使用領域と仮定）
console.log(
  UsedRange.Row,
  UsedRange.RowEnd,
  UsedRange.Column,
  UsedRange.ColumnEnd
) // 1,10,1,10
```

## Visible

ワークシートの表示／非表示を切り替えます。

#### データ型

Boolean - シートを表示するかどうか

#### 例

```js
const sheet = Application.ActiveSheet
// シートを非表示
sheet.Visible = false
// 再度表示
sheet.Visible = true
```

## Type

ワークシートの種類です。

#### データ型

[Enum.XlSheetType](Enum.md#xlsheettype) - ワークシートの種類

#### 例

```js
const sheet = Application.ActiveSheet
// アクティブなワークシートの種類を出力
console.log(sheet.Type) //xlWorksheet
```

## AutoFilter

現在のシートの AutoFilter。すべての列フィルターと、フィルターの適用／解除のためのメソッドです。

#### データ型

[AutoFilter](AutoFilter.md) - AutoFilter オブジェクト

#### 例

```js
// AutoFilter を取得
const autoFilter = Application.ActiveSheet.AutoFilter
// フィルター コレクションを取得
const filters = autoFilter.Filters
// 現在のシートのフィルター数を数える
const count = filters.Count
console.log(count)
```

## Hyperlinks

シート上のすべてのハイパーリンクのコレクションです。

#### データ型

[Hyperlinks](#hyperlinks) コレクション - すべてのハイパーリンク オブジェクト

#### 例

```js
// ハイパーリンクの数を出力
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### Hyperlinks.Count

コレクション内のハイパーリンクの数です。

#### データ型

Number - コレクション内のハイパーリンクの数

#### 例

```js
// ハイパーリンクの数を出力
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### Hyperlinks.Item()

ハイパーリンク オブジェクトを取得します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Index | Number |  | はい | 1 から始まります |

#### 戻り値の型

[Hyperlink](Hyperlink.md) - Hyperlink オブジェクト

#### 例

```js
const hyperlinks = Application.ActiveSheet.Hyperlinks
if (hyperlinks.Count <= 0) {
  throw new Error('ドキュメントにハイパーリンクがありません')
}
// ハイパーリンクの表示テキストとアドレスを出力
console.log(hyperlinks.Item(1).TextToDisplay) // ハイパーリンクの表示テキスト
console.log(hyperlinks.Item(1).Address) // ハイパーリンクの URL
```

## Shapes

現在のシート上のすべての Shape オブジェクトのコレクションです。

#### データ型

[Shapes](#shapes) コレクション - このシート上のすべての Shape オブジェクト

#### 例

```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
const chartEnum = Application.Enum.XlChartType.xlColumnClustered // 集合縦棒
// 300×300 の集合縦棒チャートを追加
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### Shapes.GetActiveShapeImg()

アクティブなセルの Shape の画像データを取得します。

#### 戻り値の型

String - 元画像のダウンロード URL

#### 例

```js
// A1 に画像がある場合
Application.Range('A1').Select()
const imgUrl = Application.ActiveSheet.Shapes.GetActiveShapeImg()
console.log(imgUrl) // https://imageUrl または画像がない場合は undefined
```

### Shapes.Item(Index)

フリーフォーム、多角形、OLE オブジェクト、画像などの描画レイヤー オブジェクトを表します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| index | Number |  | はい | オブジェクトのインデックス |

#### 戻り値の型

[Shape](Shape.md) - 描画レイヤー オブジェクトの 1 つ（フリーフォーム、多角形、OLE、画像）

#### 例

```js
const shapes = Application.ActiveSheet.Shapes
// 先頭の Shape の ID を出力
console.log(shapes.Item(1).ID)
```

### Shapes.Count

Shape の数です。

#### データ型

Number - シート上の Shape オブジェクトの数

#### 例

```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// 現在の Shape の数を出力
console.log(shapes.Count)
```

### Shapes.AddChart2()

Shape を追加します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Style | String |  | いいえ | チャートのスタイル |
| XlChartType | Enum |  | いいえ | チャートの種類。[Enum.XlChartType](Enum.md#xlcharttype) を参照 |
| Left | Number |  | いいえ | 新しいチャートの左オフセット（px） |
| Top | Number |  | いいえ | 新しいチャートの上オフセット（px） |
| Width | Number |  | いいえ | 新しいチャートの幅（px） |
| Height | Number |  | いいえ | 新しいチャートの高さ（px） |

#### 戻り値の型

[Shape](Shape.md) - 追加された Shape

#### 例

```js
const chartEnum = Application.Enum.XlChartType.xlColumnClustered // 集合縦棒の列挙型
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// 300×300 の集合縦棒チャートを追加
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### Shapes.AddPicture()

シートに浮動画像を挿入します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| FileName | String |  | はい | 挿入する画像（Base64 の data URL または画像 URL） |
| LinkToFile | Enum |  | はい | 未サポート。`0` を使用してください |
| SaveWithDocument | Enum |  | はい | 未サポート。`0` を使用してください |
| Left | Number |  | はい | シートの左端から画像の左端までの距離 |
| Top | Number |  | はい | シートの上端から画像の上端までの距離 |
| Width | Number |  | いいえ | 画像の幅 |
| Height | Number |  | いいえ | 画像の高さ |

#### 例

```js
// Shapes コレクションを取得
const shapes = ActiveSheet.Shapes
// 浮動画像を挿入
shapes.AddPicture({
  FileName:
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII=',
  LinkToFile: 0,
  SaveWithDocument: 0,
  Left: 30, // 左からのオフセット
  Top: 30, // 上からのオフセット
  Width: 32, // 画像の幅
  Height: 32 // 画像の高さ
})
```

## Sort

Sort オブジェクトです。対象の範囲を設定してから並べ替えを適用します。

#### データ型

[Sort](Sort.md) - Sort オブジェクト

#### 例

```js
// 列 C を降順、列 D を昇順で並べ替え

// 現在の使用済み範囲
const range = ActiveSheet.UsedRange
// Sort オブジェクト
const sort = ActiveSheet.Sort
// 並べ替えフィールド
const sortFields = sort.SortFields
// 既存の並べ替えキーをクリア
sortFields.Clear()
// 列 C を値で降順
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
// 列 D を値で昇順
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
// 見出し行: xlGuess は自動、xlYes は見出しあり、xlNo は見出しなし
sort.Header = xlGuess
// 大文字／小文字の区別
sort.MatchCase = false
// 中国語の並べ替え: xlPinYin または xlStroke
sort.SortMethod = xlPinYin
// xlSortColumns と xlSortRows
sort.Orientation = xlSortColumns
// Apply の前に SetRange が必要
sort.SetRange(range)
// 並べ替えを適用
sort.Apply()
```

## Range()

セルまたは範囲を Range として扱います（Range のプロパティとメソッドを使用します）。

#### データ型

[Range](Range.md) - シート上のセルまたは範囲

#### 例

```js
const sheet = Application.ActiveSheet
// アクティブなシートの D2 のテキストを出力
console.log(sheet.Range('D2').Text) // D2

// アクティブなシートの D2 に値を設定
sheet.Range('D2').Value = 'this is D2'
```

## Cells()

シート上の 1 セルを Range として選択します。

#### 戻り値の型

[Range](Range.md) - シート上の 1 セル

#### 例

```js
// アクティブなシートの先頭セルのアドレスを出力
console.log(Application.Cells(1).Address()) // $A$1
console.log(Application.Cells(2).Address()) // $A$2
```

## Activate()

ワークシートをアクティブにします。

#### 例

```js
const sheet = Application.Sheets.Item(1)
// 先頭のシートをアクティブ化。Application.ActiveSheet.Range と Application.Range はそのシートを指す
sheet.Activate()
// 先頭のシートの A1 に設定
Application.Range('A1').Value = 'foo'
```

## Move()

ワークシートを移動します。

#### パラメーター

**2 つのパラメーターは同時に使用できません。**

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Before | number | null | いいえ | 移動先のシートの直前に置くときのシート ID。After と同時には使えません。 |
| After | number | null | いいえ | 移動先のシートの直後に置くときのシート ID。Before と同時には使えません |

#### 例

```js
// 現在のシートを 2 番目のシートの後ろに移動
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## Delete()

ワークシートを削除します。

#### 戻り値

undefined

#### 例

```js
// Sheet2 という名前のシートを削除
Application.Sheets.Item('Sheet2').Delete()
```
