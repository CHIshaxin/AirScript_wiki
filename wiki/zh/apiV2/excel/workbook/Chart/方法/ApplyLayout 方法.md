---
title: "ApplyLayout (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/ApplyLayout%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DzopEC3I
---

# ApplyLayout (方法)

应用功能区中显示的版式。

## 说明

在当前图表类型上使用版式时，会将介于 1 到 10 之间的数字应用于该图表类型。还可以将某一图表类型的版式应用于另一个图表类型。例如，可以将折线图中的可用版式应用于柱形图。版式只会添加可用于该特定图表类型的图表元素。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Layout | int4 | 必填 | 指定版式类型。版式类型由 1 到 10 的数字表示。 |
| ChartType | any | 可选 | 图表的类型。 |

## 示例

```javascript
/*本示例设置活动工作表的第一个图表版式为复合饼图的第2个版式类型。*/
function test() {
    ActiveSheet.ChartObjects(1).Chart.ApplyLayout(2, xlPieOfPie)
}
```

```javascript
/*本示例设置第一张工作表的第一个图表版式为三维面积图的第7个版式类型。*/
function test() {
    let chart = Application.Sheets.Item(1).ChartObjects(1).Chart
    chart.ApplyLayout(7, xl3DArea)
}
```
