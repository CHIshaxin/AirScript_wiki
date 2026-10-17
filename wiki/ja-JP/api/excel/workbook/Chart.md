---
title: "Chart"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Chart.html
updated: 1789700916000
hash: k0aqhbjM
---

# Chart

単一の Chart オブジェクトです。

Chart オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [ChartTitle.Text](#charttitle-text) | 文字列 | グラフのタイトル |
| [HasTitle](#hastitle) | 真偽値 | グラフタイトルを表示するかどうか |

### メソッド

| メソッド | 戻り値 | 説明 |
| --- | --- | --- |
| [SetSourceData()](#setsourcedata) | undefined | グラフのソースデータ範囲を設定する |

## ChartTitle.Text

タイトルを設定します。

### データ型

文字列 - グラフのタイトル

グラフのタイトルを設定します（実質的には書き込み専用。Chart オブジェクトで確認してください）。

### 例

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// タイトルを設定
chart.ChartTitle.Text = 'Chart title'
```

## HasTitle

軸またはグラフに表示中のタイトルがある場合は true です。

タイトルの表示/非表示を設定します（実質的には書き込み専用。Chart オブジェクトで確認してください）。

### データ型

真偽値 - グラフタイトルを表示するかどうか

### 例

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// タイトルを非表示
shape.Chart.HasTitle = false
```

## SetSourceData()

グラフのソースデータ範囲を設定します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Source | Range |  | はい | ソースデータを含む Range |
| PlotBy | Enum |  | いいえ | 行または列でプロットするか。[Enum.XlRowCol](Enum.md#xlrowcol) を参照。例: xlColumns または xlRows |

### 例

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// グラフのデータ範囲
let source = Application.Range('A1:D4')

// ソースデータを適用
chart.SetSourceData(source, Application.Enum.XlRowCol.xlRows)
```
