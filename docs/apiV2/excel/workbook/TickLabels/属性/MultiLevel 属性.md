---
title: "MultiLevel 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E5%B1%9E%E6%80%A7/MultiLevel%20%E5%B1%9E%E6%80%A7.html
---
# MultiLevel (属性)

设置轴是否为多级。可读/写 **Boolean** 类型。

## 说明

此属性的默认值为 True。

## 示例

```javascript
/*下例显示 Sheet1 上的第一个图表分类轴刻度线标签的轴是否为多级轴。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    console.log(tickLabels.MultiLevel)
}
```

```javascript
/*以下示例显示 Chart1 上分类轴刻度线标签的轴是否为多级轴。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    console.log(tickLabels.MultiLevel)
}
```
