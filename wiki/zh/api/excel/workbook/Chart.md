---
title: "Chart"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Chart.html
updated: 1789700916000
hash: CJI5vp4X
---

# Chart

单个图表对象

Chart对象的具体属性和方法请参阅下方的列表。

### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| [ChartTitle.Text](#charttitle-text) | String | 图表标题 |
| [HasTitle](#hastitle) | Boolean | 图表标题是否可见 |

### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| [SetSourceData()](#setsourcedata) | undefined | 为图表设置源数据区域 |

## ChartTitle.Text

设置标题

### 数据类型

String - 图表标题

设置图形对象的标题,该属性只写,效果需要自行观察图表对象

### 示例

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// 设置标题
chart.ChartTitle.Text = '图表标题'
```

## HasTitle

如果坐标轴或图表有可见标题，则该属性值为 true

设置图形对象的标题,该属性只写,效果需要自行观察图表对象

### 数据类型

Boolean - 图表标题是否可见

### 示例

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// 设置标题不可见
shape.Chart.HasTitle = false
```

## SetSourceData()

为指定图表设置源数据区域

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Source | Range |  | 是 | 包含源数据的区域，可用 Range 对象 |
| PlotBy | Enum |  | 否 | 指定图表类型，对应 [Enum.XlRowCol](Enum.md#xlrowcol)，可以为 xlColumns 或 xlRows |

### 示例

```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// 指定图表数据源
let source = Application.Range('A1:D4')

// 设置图表数据源
chart.SetSourceData(source, Application.Enum.XlRowCol.xlRows)
```
