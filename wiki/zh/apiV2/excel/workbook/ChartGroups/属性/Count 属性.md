---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroups/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cd7FT1KR
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表组的数量是否为1。*/
function test() {
    let chartgroups = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups()
    console.log(chartgroups.Count == 1)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个内嵌图表的图表组的数量。*/
function test() {
    let chartgroups = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups()
    console.log(chartgroups.Count)
}
```
