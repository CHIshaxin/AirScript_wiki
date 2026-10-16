---
title: "XValues (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/XValues%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: COeb_xEI
---

# XValues (属性)

返回或设置图表系列中 x 值的数组。**XValues**属性可设置为工作表区域或数值数组，但不能为二者的组合。**Variant** 类型，可读写。

## 说明

对于数据透视图，该属性为只读。

## 示例

```javascript
/*本示例将 Chart1 中系列一的 x 值设置为 Sheet1 上 B1:B5 区域的值。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).XValues = Application.Worksheets.Item("Sheet1").Range("B1:B5")
}
```

```javascript
/*本示例使用数组为 Chart1 中系列一的各点逐一赋值。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).XValues = [5.0, 6.3, 12.6, 28, 50]
}
```
