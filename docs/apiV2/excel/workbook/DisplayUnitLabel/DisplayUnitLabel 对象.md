---
title: "DisplayUnitLabel 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/DisplayUnitLabel%20%E5%AF%B9%E8%B1%A1.html
---
# DisplayUnitLabel (对象)

代表指定图表中坐标轴上的单位标志。

## 说明

在绘制大数据的图表时使用单位标志会很有用，例如：上百万或几十亿的大数据。在刻度线上使用单位标志而不直接使用大数据可以使图表更易读、易理解。

使用 **DisplayUnitLabel** 属性可返回 **[DisplayUnitLabel](./../DisplayUnitLabel/DisplayUnitLabel 对象.html)** 对象。

## 示例

```javascript
/*以下示例将 Chart1 上数值轴上的显示标签标题设置为“Millions”，然后关闭自动字体缩放。*/
function test() {
    let Axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    Axis.DisplayUnit = xlMillions
    Axis.HasDisplayUnitLabel = true
    let displayUnitLabel = Axis.DisplayUnitLabel
    displayUnitLabel.Caption = "Millions"
    displayUnitLabel.AutoScaleFont = false
}
```

```javascript
/*下例删除 Sheet1 上的第一个图表数值轴上的单位标签。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Delete()
}
```
