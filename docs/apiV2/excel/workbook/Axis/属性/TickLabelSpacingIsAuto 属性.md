---
title: "TickLabelSpacingIsAuto 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/TickLabelSpacingIsAuto%20%E5%B1%9E%E6%80%A7.html
---
# TickLabelSpacingIsAuto (属性)

返回或设置刻度标签间距是否是自动设置的。可读/写 **Boolean** 类型。

## 说明

**TickLabelSpacing** 属性返回当前刻度标签间距是否是自动设置的。

## 示例

```javascript
/*本示例自动设置图表工作表 Chart1 中图表的分类轴上刻度标签间距。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.TickLabelSpacingIsAuto = true
}
```

```javascript
/*本示例关闭工作表 Sheet1 中第一张图表的分类轴上自动设置刻度标签间距，并将该刻度线标签间距设置为 2。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.TickLabelSpacingIsAuto = false
    axis.TickLabelSpacing = 2
}
```
