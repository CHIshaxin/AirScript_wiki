---
title: "Chart 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Chart%20%E5%B1%9E%E6%80%A7.html
---
# Chart (属性)

返回 **[Chart](./../../Chart/Chart 对象.html)** 对象，该对象代表形状中包含的图表。只读。

## 示例

```javascript
/*本示例将活动工作表上第一个形状的标题设置为"销售数据"。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    let chart = shape.Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "销售数据"
}
```

```javascript
/*本示例将第一张工作表上的第一个形状的X轴标题设置为"月份"。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    let chart = shape.Chart
    chart.Axes(xlCategory).HasTitle = true
    chart.Axes(xlCategory).AxisTitle.Text = "月份"
}
```
