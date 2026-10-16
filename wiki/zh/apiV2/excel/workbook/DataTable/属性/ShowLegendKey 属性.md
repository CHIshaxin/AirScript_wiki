---
title: "ShowLegendKey (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/%E5%B1%9E%E6%80%A7/ShowLegendKey%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: WsIYopv5
---

# ShowLegendKey (属性)

如果数据标签图例项标示可见，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的模拟运算表的图例项标示设置为可见。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    datatable.ShowLegendKey = true
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的模拟运算表的图例项标示是否可见。*/
function test() {
    let datatable = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.DataTable
    console.log(datatable.ShowLegendKey)
}
```
