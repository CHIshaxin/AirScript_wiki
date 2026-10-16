---
title: "InvertIfNegative 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/InvertIfNegative%20%E5%B1%9E%E6%80%A7.html
---
# InvertIfNegative (属性)

如果指定项与一个负数相对应时 ET 就将其反色，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例反转 Chart1 中第一个数据系列的负值的图案。此示例须在二维柱形图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).InvertIfNegative = true
}
```

```javascript
/*此示例显示活动工作表中第一个图表的第二个数据系列在负值情况下是否反转颜色。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(2).InvertIfNegative)
}
```
