---
title: "IsFiltered (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/IsFiltered%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BVqXwyZM
---

# IsFiltered (属性)

此设置控制是否已从图表中筛选出序列。 默认值为 **False**。 读/写 **Boolean**。

## 说明

当用户筛选出某个系列时， **系列 IsFiltered** 属性将切换到 **True**，并且该系列将从其父 **SeriesCollection** 中移出。 但是，该系列仍保留在其父 **FullSeriesCollection** 中。 当用户重新筛选序列时，该序列将插入到 **SeriesCollection** 中的原始位置。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 上的第三个数据系列是否被筛选出序列。*/
function test() {
    let series3 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    console.log(series3.IsFiltered)
}
```

```javascript
/*此示例将第一张工作表上的第一个图表的第二个数据系列筛选出序列。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series2.IsFiltered = true
}
```
