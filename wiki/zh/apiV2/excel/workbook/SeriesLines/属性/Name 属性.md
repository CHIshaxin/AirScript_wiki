---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesLines/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: VoGnoW9d
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 上第一个图表组的系列线名称。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.SeriesLines.Name)
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表上第一个图表组的系列线名称。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.SeriesLines.Name)
}
```
