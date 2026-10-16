---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DownBars/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CSNcAnU8
---

# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*本示例选中图表工作表 Chart1 中图表的第一个图表组的跌柱线。*/
function test() {
    let downbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DownBars
    downbars.Select()
}
```

```javascript
/*本示例选中工作表 Sheet1 中第一个图表的第一个图表组的跌柱线，并设置该跌柱线的前景色为红色。*/
function test() {
    let downbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).DownBars
    downbars.Select()
    Selection.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```
