---
title: "MinimumScaleIsAuto 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MinimumScaleIsAuto%20%E5%B1%9E%E6%80%A7.html
---
# MinimumScaleIsAuto (属性)

如果 ET 为数值轴计算最小值，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

设置 **MinimumScale** 属性会将该属性设置为 False。

## 示例

```javascript
/*本示例自动计算图表工作表 Chart1 中的数值轴的最小刻度和最大刻度。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MinimumScaleIsAuto = true
    axis.MaximumScaleIsAuto = true
}
```

```javascript
/*本示例关闭工作表 Sheet1 中的数值轴的最小刻度自动计算，并将数值轴的最小刻度设置为 2.5。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MinimumScaleIsAuto = false
    axis.MinimumScale = 2.5
}
```
