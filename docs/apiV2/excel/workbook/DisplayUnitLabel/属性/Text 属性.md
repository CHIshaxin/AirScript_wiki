---
title: "Text 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E5%B1%9E%E6%80%A7/Text%20%E5%B1%9E%E6%80%A7.html
---
# Text (属性)

返回或设置指定对象中的文本。**String** 型，可读写。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中数值轴上的单位标签文本。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Text)
}
```

```javascript
/*下例设置 Sheet1 上的第一个图表数值轴上的单位标签文本。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Text = "万"
}
```
