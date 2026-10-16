---
title: "Chart"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Chart.html
updated: 1789700916000
hash: Cu9-7RPX
---

# Chart

Single chart object.

See the lists below for the Chart object's properties and methods.

### Properties

| Property | Data type | Description |
|---|---|---|
| [ChartTitle.Text](#charttitle-text) | String | Chart title |
| [HasTitle](#hastitle) | Boolean | Whether chart title is visible |

### Methods

| Method | Return type | Description |
|---|---|---|
| [SetSourceData()](#setsourcedata) | undefined | Sets the chart source data range |

## ChartTitle.Text

Sets the title.

### Data type

String - Chart title

Sets the chart title (write-only in practice; observe the chart object).

### Example

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// Set title
chart.ChartTitle.Text = 'Chart title'
```

## HasTitle

True if the axis or chart has a visible title.

Sets title visibility (write-only in practice; observe the chart object).

### Data type

Boolean - Whether chart title is visible

### Example

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// Hide title
shape.Chart.HasTitle = false
```

## SetSourceData()

Sets the source data range for the chart.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Source | Range |  | Yes | Range containing source data |
| PlotBy | Enum |  | No | Plot by rows or columns; see [Enum.XlRowCol](/docs/en-US/api/excel/workbook/Enum.html#xlrowcol), e.g. xlColumns or xlRows |

### Example

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// Chart data range
let source = Application.Range('A1:D4')

// Apply source data
chart.SetSourceData(source, Application.Enum.XlRowCol.xlRows)
```
