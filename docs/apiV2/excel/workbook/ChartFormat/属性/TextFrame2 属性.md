---
title: "TextFrame2 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/TextFrame2%20%E5%B1%9E%E6%80%A7.html
---
# TextFrame2 (属性)

返回一个包含指定图表元素文本格式的 **[TextFrame2](./../../TextFrame2/TextFrame2 对象.html)** 对象。只读。

## 示例

```javascript
/*此示例设置图表工作表 Chart1 图表区文本框中字体的大小为 10 磅。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.TextFrame2.TextRange.Font.Size = 10
}
```

```javascript
/*此示例显示工作表 Sheet1 第一个内嵌图表区的文本框方向是否为水平方向。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    console.log(chartformat.TextFrame2.Orientation == msoTextOrientationHorizontal)
}
```
