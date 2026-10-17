---
title: "Font"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Font.html
updated: 1789700916000
hash: Bd3c3eg5
---

# Font

セル内テキストの Font プロパティ。太字、色、サイズ、斜体、取り消し線、下線などです。

Font オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [Bold](#bold) | 真偽値 | 太字かどうか |
| [Color](#color) | 文字列 | フォントの色 |
| [Italic](#italic) | 真偽値 | 斜体かどうか |
| [Name](#name) | 文字列 | フォント名 |
| [Size](#size) | 数値 | フォントサイズ |
| [Strikethrough](#strikethrough) | 真偽値 | 取り消し線を付けるかどうか |
| [Underline](#underline) | [Enum.XlUnderlineStyle](Enum.md#xlunderlinestyle) | 下線のスタイル |

## Bold

太字の書式を取得または設定します。

### データ型

真偽値 - 太字かどうか

### 例

```js
// Font オブジェクト
let font = Application.Range('A1').Font

// 太字かどうかを出力
console.log('Bold:', font.Bold)

// 太字に設定
font.Bold = true
```

## Color

フォントの色を取得または設定します。

### データ型

文字列 - 16 進 RGB によるフォントの色

### 例

```js
// Font オブジェクト
let font = Application.Range('A1').Font

// フォントの色を出力
console.log('Color:', font.Color)

// フォントの色を設定
font.Color = '#eb5451'
```

## Italic

斜体の書式を取得または設定します。

### データ型

真偽値 - 斜体かどうか

### 例

```js
// Font オブジェクト
let font = Application.Range('A1').Font

// 斜体かどうかを出力
console.log('Italic:', font.Italic) // Italic: false

// 斜体に設定
font.Italic = true
```

## Name

フォント名を取得または設定します。

### データ型

文字列 - フォント名

### 例

```js
// A1 のフォントを Microsoft YaHei に設定
Range('A1').Font.Name = 'Microsoft YaHei'
```

## Size

フォントサイズを取得または設定します。

### データ型

数値 - フォントサイズ

### 例

```js
// Font オブジェクト
let font = Application.Range('A1').Font

// フォントサイズを出力
console.log(font.Size) //12

// フォントサイズを設定
font.Size = 30
```

## Strikethrough

取り消し線を取得または設定します。

### データ型

真偽値 - 取り消し線を付けるかどうか

### 例

```js
// Font オブジェクト
let font = Application.Range('A1').Font

// 取り消し線のオン/オフを出力
console.log('Strikethrough:', font.Strikethrough) // Strikethrough: false

// 取り消し線を有効化
font.Strikethrough = true
```

## Underline

下線スタイルを取得または設定します。

### データ型

[Enum.XlUnderlineStyle](Enum.md#xlunderlinestyle) - 下線の種類

例: 単線の下線は `Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle`

### 例

```js
// Font オブジェクト
const font = Application.Range('A1').Font

// 下線の設定を出力
console.log('Underline:', font.Underline) // Underline: -4142

const underlineStatus = Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle

// 単線の下線
font.Underline = underlineStatus
```
