---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BFM6C9Bl
---

# Orientation (属性)

返回或设置一个 **Variant** 值，它代表文本方向。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的分类轴标题的文字方向是否为水平排列。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    console.log(axistitle.Orientation == xlHorizontal)
}
```

```javascript
/*本示例判断如果工作表 Sheet1 中第一张图表的数值轴标题为可见，则将该标题的文字方向设置为向下排列。*/
function test() {
    let axis = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    if (axis.HasTitle) {
        axis.AxisTitle.Orientation = xlDownward
    }
}
```
