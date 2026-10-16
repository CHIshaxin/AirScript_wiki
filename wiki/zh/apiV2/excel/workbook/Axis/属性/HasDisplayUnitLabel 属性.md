---
title: "HasDisplayUnitLabel (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/HasDisplayUnitLabel%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cr_Y6Vwf
---

# HasDisplayUnitLabel (属性)

如果由 **DisplayUnit** 或 **DisplayUnitCustom** 属性指定的标签显示在指定坐标轴上，则该属性值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中数值轴上的单位设置为 500，但隐藏单位标志。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.DisplayUnit = xlCustom
    axis.DisplayUnitCustom = 500
    axis.HasTitle = true
    axis.AxisTitle.Caption = "Rebate Amounts"
    axis.HasDisplayUnitLabel = false
}
```

```javascript
/*本示例将工作表 Sheet1 中第一个图表的数值轴上的单位设置为 10，并显示单位标志是否可见。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.DisplayUnit = xlCustom
    axis.DisplayUnitCustom = 10
    console.log(axis.HasDisplayUnitLabel)
}
```
