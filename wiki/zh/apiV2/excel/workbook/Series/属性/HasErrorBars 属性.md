---
title: "HasErrorBars (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/HasErrorBars%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CE1G5Yin
---

# HasErrorBars (属性)

如果数据系列有误差线，该值为 **True**。该属性对三维图表无效。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例删除 Chart1 中第一个数据系列的误差线。本示例应在其第一个系列有误差线的二维折线图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).HasErrorBars = false
}
```

```javascript
/*此示例显示活动工作表中第一个图表的第三个数据系列是否有误差线。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(3).HasErrorBars)
}
```
