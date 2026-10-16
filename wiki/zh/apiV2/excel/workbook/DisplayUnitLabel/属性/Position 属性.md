---
title: "Position (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E5%B1%9E%E6%80%A7/Position%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DBXT83IS
---

# Position (属性)

返回或设置图表上单位标签的位置。可读/写 **XlChartElementPosition** 类型。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中数值轴上的单位标签位置是否为自动设置。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Position == xlChartElementPositionAutomatic)
}
```

```javascript
/*下例为 Sheet1 上的第一个图表数值轴上的单位标签设置指定特定位置。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Position = xlChartElementPositionCustom
}
```
