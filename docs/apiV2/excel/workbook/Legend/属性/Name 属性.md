---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示第一张工作表上的第一个图表中图例的名称。*/
function test() {
    console.log(Worksheets.Item(1).ChartObjects(1).Chart.Legend.Name)
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第一个图表中图例的名称。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Name)
}
```
