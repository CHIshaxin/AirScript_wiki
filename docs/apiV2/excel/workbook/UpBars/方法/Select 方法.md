---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UpBars/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*本示例选中图表工作表 Chart1 中图表的第一个图表组的涨柱线。*/
function test() {
    let upbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    upbars.Select()
}
```

```javascript
/*本示例选中工作表 Sheet1 中第一个图表的第一个图表组的涨柱线，并显示该涨柱线的名称。*/
function test() {
    let upbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    upbars.Select()
    console.log(Selection.Name)
}
```
