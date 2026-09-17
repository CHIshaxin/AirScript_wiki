---
title: "LogBase 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/LogBase%20%E5%B1%9E%E6%80%A7.html
---
# LogBase (属性)

在使用对数刻度时返回或设置对数的底。可读/写 **Double** 类型。

## 说明

试图将此属性设置为小于或等于零 (0) 的值将导致错误。默认值是 10。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值轴单位的对数的底。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes().Item(xlValue)
    console.log(axis.LogBase)
}
```

```javascript
/*本示例将图表工作表 Chart1 中的数值轴的单位设置为 400，并将其刻度对数的底设置为 8。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlValue)
    axis.DisplayUnit = xlCustom
    axis.DisplayUnitCustom = 400
    axis.LogBase = 8
}
```
