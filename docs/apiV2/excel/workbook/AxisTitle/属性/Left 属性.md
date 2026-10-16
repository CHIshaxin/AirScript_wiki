---
title: "Left 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
---
# Left (属性)

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一张图表的分类轴标题设置为距离图表区左边缘 20 磅。*/
function test() {
    let axistitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.Left = 20
}
```

```javascript
/*本示例显示图表工作表 Chart1 中图表的数值轴标题到图表区左边缘的距离。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).AxisTitle
    console.log(`数值轴标题到图表左边缘的距离：${axistitle.Left} 磅`)
}
```
