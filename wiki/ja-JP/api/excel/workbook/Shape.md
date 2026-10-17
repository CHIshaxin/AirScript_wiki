---
title: "Shape"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Shape.html
updated: 1789700916000
hash: gHuHCjyn
---

# Shape

ワークシート上の単一の Shape グラフィック オブジェクトです。

Shape オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [ID](#id) | String | Shape の ID |
| [Name](#name) | String | Shape の名前 |
| [Title](#title) | String | Shape のタイトル |
| [Chart](#chart) | [Chart](Chart.md) | Chart オブジェクト |

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [Delete](#delete) | undefined | Shape を削除します |

## ID

Shape の ID です。

### データ型

String - Shape オブジェクトの ID

### 例

```js
// Shape コレクション
const shape = Application.ActiveSheet.Shapes.Item(1)

// Shape の ID を出力
console.log(shape.ID)  //94646582190480
```

## Name

Shape の名前です。

### データ型

String - Shape オブジェクトの名前

### 例

```js
// Shape コレクション
const shape = Application.ActiveSheet.Shapes.Item(1)

// Shape の名前を出力
console.log(shape.Name) // Chart 1
```

## Title

Shape のタイトルです。

書き込み専用で、読み取りはできません。

### データ型

String - Shape のタイトル

### 例

```js
const shape = Application.ActiveSheet.Shapes.Item(1)

// タイトルを設定（書き込み専用。チャート上の表示で確認）
shape.Title = 'WebOffice'
```

## Chart

単一の Chart オブジェクトです。

### データ型

[Chart](Chart.md) - Chart オブジェクト

### 例

```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Chart.HasTitle = false // タイトルを非表示
```

## Delete()

Shape を削除します。

### データ型

undefined

### 例

```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Delete()
```
