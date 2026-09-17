---
title: "Position 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E5%B1%9E%E6%80%A7/Position%20%E5%B1%9E%E6%80%A7.html
---
# Position (属性)

返回或设置一个 **XlLegendPosition**值，它代表图表上图例的位置。

## 示例

```javascript
/*本示例将图表中的图例移到图表的底部。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Position = xlLegendPositionBottom
}
```

```javascript
/*本示例显示第一张工作表上的第一个图表中图例的位置是否位于图表上方。*/
function test() {
    console.log(Worksheets.Item(1).ChartObjects(1).Chart.Legend.Position == xlLegendPositionTop)
}
```
