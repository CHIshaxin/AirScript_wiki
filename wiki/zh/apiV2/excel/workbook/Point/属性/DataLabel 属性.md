---
title: "DataLabel (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/DataLabel%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CxN7Qx8r
---

# DataLabel (属性)

返回一个 **[DataLabel](./../../DataLabel/DataLabel 对象.html)** 对象，它代表与数据点相关的数据标签。只读。

## 示例

```javascript
/*此示例打开 Chart1 上第三个数据系列的第七个数据点的数据标签，并将该数据标签的颜色设置为蓝色。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Points(7)
    point.HasDataLabel = true
    point.ApplyDataLabels(xlValue)
    point.DataLabel.Font.ColorIndex = 5
}
```

```javascript
/*下例打开Chart1的图表工作表上第二个系列中第三个数据点的数据标签，并将数据标签文本设置为Saturday。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Points(3)
    point.HasDataLabel = true
    point.DataLabel.Text = "Saturday"
}
```
