---
title: "Border"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Border.html
updated: 1789700916000
hash: QwMDdual
---

# Border

Border オブジェクト。Borders コレクション内の 1 つの罫線です。

Border オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [Color](#color) | 数値 | 罫線の色 |
| [Weight](#weight) | [Enum.XlBorderWeight](/docs/ja-JP/api/excel/workbook/Enum.html#xlborderweight) | 罫線の太さ |
| [LineStyle](#linestyle) | [Enum.XlLineStyle](/docs/ja-JP/api/excel/workbook/Enum.html#xllinestyle) | 罫線の線種 |

## Color

罫線の色です。

備考: 罫線の色を読み取るときは、具体的な罫線を指定してください。[Enum.XlBordersIndex](/docs/ja-JP/api/excel/workbook/Enum.html#xlbordersindex) に xlAll、xlOutside、xlInside などを指定してはいけません。

### データ型

数値 - 罫線の色

### 例

```js
// Range の Borders
const borders = Application.Range('A1').Borders

// 下罫線
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// 罫線の色を出力
console.log(border.Color) // {"private":{"isAuto":false,"rgbValue":"#000000","tint":0,"type":"ectAUTO"},"rgbValue":"#000000"}

// 外枠をすべて緑に設定
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.Color = '#00FF00'
```

## Weight

罫線の太さです。

### データ型

[Enum.XlBorderWeight](/docs/ja-JP/api/excel/workbook/Enum.html#xlborderweight) - 罫線の太さ

書き込み専用。読み取れません。

### 例

```js
// Range の Borders
let borders = Application.Range('A3').Borders

// 単一の罫線
let border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)

// 太さを設定
border.Weight = Application.Enum.XlBorderWeight.xlThick
```

## LineStyle

罫線の線種です。

### データ型

[Enum.XlLineStyle](/docs/ja-JP/api/excel/workbook/Enum.html#xllinestyle) - 線種

書き込み専用。読み取れません。

### 例

```js
// Range の Borders
let borders = Application.Range('A1').Borders

// 下罫線
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// 外枠を二重線に設定
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.LineStyle = Application.Enum.XlLineStyle.xlDouble
```
