---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/HiLoLines/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组高低点连线的名称。*/
function test() {
    let chartGroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartGroup.HiLoLines.Name)
}
```

```javascript
/*本示例将工作表 Sheet1 上第一个图表中第一个图表组高低点连线的名称赋值给C5单元格。*/
function test() {
    let chartGroup = Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    Range("C5").Value2 = chartGroup.HiLoLines.Name
}
```
