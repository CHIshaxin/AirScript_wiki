---
title: "Height 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
---
# Height (属性)

返回或设置一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*本示例将第一张工作表上的第一个图表的图例的高度设置为15磅。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.Legend.Height = 15
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第一个图表中的图例高度。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Height)
}
```
