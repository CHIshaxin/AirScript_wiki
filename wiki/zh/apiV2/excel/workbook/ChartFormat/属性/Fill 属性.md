---
title: "Fill (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/Fill%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DKR74dMO
---

# Fill (属性)

返回一个包含图表元素填充格式属性的父图表元素的 **[FillFormat](./../../FillFormat/FillFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*此示例设置图表工作表 Chart1 图表区的颜色为黄色。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.Fill.ForeColor.RGB = RGB(255, 255, 0)
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一个内嵌图表区的填充渐变样式是否为混合的。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    console.log(chartformat.Fill.GradientStyle == msoGradientMixed)
}
```
