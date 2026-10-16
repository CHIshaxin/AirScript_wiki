---
title: "MinimumScale (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MinimumScale%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DeoGUSUR
---

# MinimumScale (属性)

返回或设置数值轴上的最小值。**Double** 类型，可读写。

## 说明

设置该属性会将 **MinimumScaleIsAuto** 属性设置为 **False**。

## 示例

```javascript
/*本示例对图表工作表 Chart1 中的数值轴的最小值和最大值进行设置。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MinimumScale = 10
    axis.MaximumScale = 120
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴的最小值设置为该工作表单元格 E1 中的值。*/
function test() {
    let sheet = Application.Worksheets.Item("Sheet1")
    let axis = sheet.ChartObjects(1).Chart.Axes(xlValue)
    axis.MinimumScale = sheet.Range("E1").Value2
}
```
