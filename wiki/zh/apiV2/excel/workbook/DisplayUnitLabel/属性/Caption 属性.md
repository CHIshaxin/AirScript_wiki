---
title: "Caption (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E5%B1%9E%E6%80%A7/Caption%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: N-GvE-sY
---

# Caption (属性)

返回或设置一个 **String** 值，它代表显示单位标志文本。

## 示例

```javascript
/*以下示例将 Chart1 上数值轴上显示标签的标题设置为“Millions”。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Caption = "Millions"
}
```

```javascript
/*下例显示 Sheet1 上的第一个图表数值轴上单位标签的标题。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Caption)
}
```
