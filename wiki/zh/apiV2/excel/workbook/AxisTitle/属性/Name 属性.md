---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D3P9tYYO
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的数值轴标题对象的名称。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).AxisTitle
    console.log(axistitle.Name)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的分类轴标题对象的名称。*/
function test() {
    let axistitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    console.log(axistitle.Name)
}
```
