---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Be_YwAGb
---

# Height (属性)

返回对象的高度（以磅为单位）。只读。

## 返回值

Double

## 示例

```javascript
/*下例显示 Sheet1 上的第一个图表单位标签高度。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Height)
}
```

```javascript
/*本示例显示 Chart1 上单位标签高度。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Height)
}
```
