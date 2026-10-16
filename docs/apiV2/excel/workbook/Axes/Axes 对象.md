---
title: "Axes 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axes/Axes%20%E5%AF%B9%E8%B1%A1.html
---
# Axes (对象)

指定图表中所有 **[Axis](./../Axis/Axis 对象.html)** 对象的集合。

## 示例

```javascript
/*本示例为图表工作表 Chart1 中分类轴设置标题文本。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlValue)
    axis.AxisTitle.Caption = "成绩"
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的坐标轴的数量是否为2。*/
function test() {
    let axes = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes()
    console.log(axes.Count == 2)
}
```
