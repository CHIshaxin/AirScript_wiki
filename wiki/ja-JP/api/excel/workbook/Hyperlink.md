---
title: "Hyperlink"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Hyperlink.html
updated: 1789700916000
hash: z6MToUyj
---

# Hyperlink

単一の Hyperlink オブジェクトです。ジャンプ先のアドレスと表示テキストがあり、両者は一致している必要はありません。

Hyperlink オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [Address](#address) | 文字列 | ハイパーリンクのリンク先 URL |
| [TextToDisplay](#texttodisplay) | 文字列 | ハイパーリンクの表示テキスト |

## Address

ハイパーリンクのアドレスを取得または設定します。

### データ型

文字列 - ハイパーリンクのリンク先 URL。

### 例

```js
// Hyperlink オブジェクトを取得
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// ハイパーリンクのアドレスを出力
console.log(hyperlink.Address) // https://www.wps.com
hyperlink.Address = "https://www.wps.com"
```

## TextToDisplay

ハイパーリンクの表示テキストを取得または設定します。

### データ型

文字列 - ハイパーリンクの表示テキスト。

### 例

```js
// Hyperlink オブジェクトを取得
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// 表示テキストを出力
console.log(hyperlink.TextToDisplay) // WPS Docs のサイト
hyperlink.TextToDisplay = "WPS Docs"
```
