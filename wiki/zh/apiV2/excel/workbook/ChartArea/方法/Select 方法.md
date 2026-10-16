---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DS4myOQZ
---

# Select (方法)

选择对象。

## 示例

```javascript
/*此示例选中图表工作表 Chart1 的图表区。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Select()
}
```

```javascript
/*此示例选中 Sheet1 工作表中第一个内嵌图表的图表区。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    chartarea.Select()
}
```
