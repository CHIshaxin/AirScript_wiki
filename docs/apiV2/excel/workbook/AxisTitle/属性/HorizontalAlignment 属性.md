---
title: "HorizontalAlignment 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/HorizontalAlignment%20%E5%B1%9E%E6%80%A7.html
---
# HorizontalAlignment (属性)

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

## 说明

此属性的值可以设置为 **XlHAlign** 常量之一。

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一张图表的分类轴标题的水平对齐方式设置为靠右。*/
function test() {
    let axistitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.HorizontalAlignment = xlHAlignRight
}
```

```javascript
/*本示例将图表工作表 Chart1 中图表的数值轴标题的水平对齐方式设置为居中。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).AxisTitle
    axistitle.HorizontalAlignment = xlHAlignCenter
}
```
