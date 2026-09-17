---
title: "DisplayUnitCustom 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/DisplayUnitCustom%20%E5%B1%9E%E6%80%A7.html
---
# DisplayUnitCustom (属性)

如果 **DisplayUnit** 属性的值是 **xlCustom**，则 **DisplayUnitCustom** 返回或设置显示的单位的值。该值范围必须是从 0 到 10E307。**Double** 类型，可读写。

## 说明

在绘制大数值图表时，使用单位标志可使刻度线标志易于阅读。例如，如果将数值轴的单位设置为百、千或百万，则可以在坐标轴的刻度标志上使用较小的数字值。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表数值轴上显示的单位设置为 500。*/
function test() {
    let axes = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axes.DisplayUnit = xlCustom
    axes.DisplayUnitCustom = 500
    axes.HasTitle = true
    axes.AxisTitle.Caption = "Rebate Amounts"
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值轴上的单位。*/
function test() {
    let axes = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axes.DisplayUnitCustom)
}
```
