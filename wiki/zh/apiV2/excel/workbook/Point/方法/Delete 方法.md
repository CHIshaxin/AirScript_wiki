---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cm7TJRqL
---

# Delete (方法)

删除数据点所属的系列。

## 说明

删除 **[Point](./../../Point/Point 对象.html)** 对象将删除整个数据系列。

## 返回值

Variant

## 示例

```javascript
/*此示例将删除活动工作表第一个图表上第三个系列。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(3).Points(3).Delete()
}
```

```javascript
/*下例判断Chart1的图表工作表上第二个系列中第三个数据点的数据标签如果已打开，则删除该数据点所属系列，否则打开该数据点的数据标签。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Points(3)
    if (point.HasDataLabel) {
        point.Delete()
        console.log("数据点所属系列已删除")
    } else {
        point.HasDataLabel = true
        console.log("数据点的数据标签已打开")
    }
}
```
