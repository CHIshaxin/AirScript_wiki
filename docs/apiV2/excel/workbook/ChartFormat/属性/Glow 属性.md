---
title: "Glow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/Glow%20%E5%B1%9E%E6%80%A7.html
---
# Glow (属性)

返回一个包含图表元素发光格式属性的指定图表的 **GlowFormat** 对象。只读。

## 说明

发光效果为图形添加了色彩鲜明的彩色边缘。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 图表区发光的透明度。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    console.log(chartformat.Glow.Transparency)
}
```

```javascript
/*此示例设置工作表 Sheet1 中第一个内嵌图表区的发光效果的半径值为 1 磅。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.Glow.Radius = 1
}
```
