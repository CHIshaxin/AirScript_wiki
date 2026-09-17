---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*此示例设置工作表 Sheet1 中第一张图表的宽度为 360 磅。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Width = 360
}
```

```javascript
/*此示例显示第一张工作表中图表 Chart 2 的宽度。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects("Chart 2")
    console.log(`图表 Chart 2 的宽度：${chartobject.Width} 磅`)
}
```
