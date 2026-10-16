---
title: "Smooth (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/Smooth%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BVNgbKk7
---

# Smooth (属性)

如果折线图或散点图有平滑线，则为 **True**。仅适用于折线图和散点图。可读写。

## 示例

```javascript
/*此示例使 Chart1 中的系列一具有平滑线。此示例应在二维折线图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Smooth = true
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列是否启用曲线平滑处理。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.Smooth)
}
```
