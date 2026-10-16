---
title: "ApplyPictToEnd (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/ApplyPictToEnd%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bp7QjXoo
---

# ApplyPictToEnd (属性)

如果图片应用于系列中数据点或所有数据点之后，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例将图片应用于第一个数据系列中所有数据点之后。图片必须已应用于该系列（设置此属性将更改图片的方向）。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).ApplyPictToEnd = true
}
```

```javascript
/*此示例显示图片是否应用于活动工作表上的第一个图表的第二个系列的数据点之后。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(2).ApplyPictToEnd)
}
```
