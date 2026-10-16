---
title: "HasDataLabels 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/HasDataLabels%20%E5%B1%9E%E6%80%A7.html
---
# HasDataLabels (属性)

如果数据系列具有数据标签，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例显示 Chart1 中第三个数据系列的数据标签。*/
function test() {
    let Ser = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    Ser.HasDataLabels = true
    Ser.ApplyDataLabels(xlValue)
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第二个数据系列是否有数据标签。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.HasDataLabels)
}
```
