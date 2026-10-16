---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Gridlines/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BQFW1Wmr
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 数值轴的主要网格线的名称。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axis.MajorGridlines.Name)
}
```

```javascript
/*本示例将工作表 Sheet1 中第一个图表类别轴的主要网格线名称赋值给A5单元格。*/
function test() {
    let axis = Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    Range("A5").Value2 = axis.MajorGridlines.Name
}
```
