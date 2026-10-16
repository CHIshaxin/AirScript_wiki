---
title: "MaximumScaleIsAuto (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MaximumScaleIsAuto%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: rKntNM3x
---

# MaximumScaleIsAuto (属性)

如果 ET 计算该数值轴的最大值，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

设置 **MaximumScale** 属性会将该属性设置为 False。

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
/*本示例判断如果工作表 Sheet1 中第一张图表的数值轴未开启自动计算最小刻度，则开启该数值轴自动计算最大刻度。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    if (axis.MinimumScaleIsAuto == false) {
        axis.MaximumScaleIsAuto = true
    }
}
```
