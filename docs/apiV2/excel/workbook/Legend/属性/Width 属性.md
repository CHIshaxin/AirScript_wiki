---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*本示例将第一张工作表上的第一个图表的图例的宽度设置为200磅。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.Legend.Width = 200
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第一个图表中的图例宽度。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Width)
}
```
