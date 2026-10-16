---
title: "HasDataLabel 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/HasDataLabel%20%E5%B1%9E%E6%80%A7.html
---
# HasDataLabel (属性)

如果数据点具有数据标签，则该属性值为 **True**。**Boolean** 类型，可读写。

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
/*此示例显示活动工作表第一个图表上第二个数据点是否打开数据标签。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    console.log(ActiveChart.SeriesCollection(1).Points(2).HasDataLabel)
}
```
