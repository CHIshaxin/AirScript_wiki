---
title: "Left (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DFgo0ulL
---

# Left (属性)

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例将第一张工作表上的第一个图表的图例的左侧边缘位置设置为 10 磅。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.Legend.Left = 10
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第一个图表中图例的左侧边缘位置。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Left)
}
```
