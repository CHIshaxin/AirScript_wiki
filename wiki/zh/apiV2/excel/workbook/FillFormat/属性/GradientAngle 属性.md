---
title: "GradientAngle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/GradientAngle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: NwyH8rZx
---

# GradientAngle (属性)

返回或设置指定填充格式的渐变填充角度。可读/写。

## 说明

可以在图表中各种元素（形状）的格式设置中指定渐变填充。例如，可使用**“设置数据系列格式”**对话框将**“列”**图表中各列的格式设置为渐变填充。在此情况下，**GradientAngle** 属性对应于**“设置数据系列格式”**对话框内**“填充”**类别中的**“角度”**框的设置。**GradientAngle**属性值的有效范围是 0 到 359.9。

## 示例

```javascript
/*下面的代码示例将 Chart 1 中第 1 个系列的渐变填充角度设置为 45 度。*/
function test() {
    Application.ActiveSheet.ChartObjects("Chart 1").Activate()
    Application.ActiveChart.SeriesCollection(1).Format.Fill.GradientAngle = 45
}
```

```javascript
/*本示例显示第一张工作表中第二个形状的填充格式的渐变填充的角度。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(2).Fill
    console.log(fillFormat.GradientAngle)
}
```
