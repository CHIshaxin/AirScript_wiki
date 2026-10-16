---
title: "MinorUnit 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MinorUnit%20%E5%B1%9E%E6%80%A7.html
---
# MinorUnit (属性)

返回或设置数值轴上的次要单位。**Double** 类型，可读写。

## 说明

设置该属性会将 **MinorUnitIsAuto** 属性设置为 **False**。 使用 **TickMarkSpacing** 属性可设置分类轴上的刻度线间距。

## 示例

```javascript
/*本示例为图表工作表 Chart1 中的数值轴设置主要单位和次要单位。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MajorUnit = 100
    axis.MinorUnit = 20
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值轴的次要单位。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(`数值轴上的次要单位：${axis.MinorUnit}`)
}
```
