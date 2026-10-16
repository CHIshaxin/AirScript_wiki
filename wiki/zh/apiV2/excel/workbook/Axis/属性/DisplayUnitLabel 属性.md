---
title: "DisplayUnitLabel (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/DisplayUnitLabel%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: aPrRyaBA
---

# DisplayUnitLabel (属性)

返回指定坐标轴的 **[DisplayUnitLabel](./../../DisplayUnitLabel/DisplayUnitLabel 对象.html)** 对象。如果 **HasDisplayUnitLabel** 属性设置为 **False**，则返回 **null**。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中数值轴的标签标题设置为 “Millions” ，然后关闭自动字体缩放。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.DisplayUnitLabel.Caption = "Millions"
    axis.DisplayUnitLabel.AutoScaleFont = false
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表数值轴上显示的单位设置为 200 ，并将单位标签设置为加粗。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.DisplayUnit = xlCustom
    axis.DisplayUnitCustom = 200
    axis.DisplayUnitLabel.Characters().Font.Bold = true
}
```
