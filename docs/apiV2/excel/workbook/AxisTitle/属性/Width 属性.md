---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

返回对象的宽度（以磅为单位）。只读。

## 返回值

Double

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的分类轴标题的宽度。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    console.log(`分类轴标题的宽度为 ${axistitle.Width} 磅`)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值轴标题的宽度是否为 55.6 磅。*/
function test() {
    let axistitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).AxisTitle
    console.log(axistitle.Width == 55.6)
}
```
