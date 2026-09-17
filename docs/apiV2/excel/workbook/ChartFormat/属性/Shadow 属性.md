---
title: "Shadow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
---
# Shadow (属性)

返回一个 **[ShadowFormat](./../../ShadowFormat/ShadowFormat 对象.html)** 对象，该对象包含图表元素的阴影格式属性。只读。

## 示例

```javascript
/*此示例显示工作表 Sheet1 第一个内嵌图表区的底纹的模糊度。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    console.log(chartformat.Shadow.Blur)
}
```

```javascript
/*此示例设置图表工作表 Chart1 图表区的阴影类型。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.Shadow.Type = msoShadow8
}
```
