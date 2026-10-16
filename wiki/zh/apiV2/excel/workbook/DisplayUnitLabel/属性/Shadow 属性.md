---
title: "Shadow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: pEU4zJyh
---

# Shadow (属性)

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中数值轴上的单位标签是否有阴影。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Shadow)
}
```

```javascript
/*下例将 Sheet1 上的第一个图表数值轴上的单位标签设置成有阴影。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Shadow = true
}
```
