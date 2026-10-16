---
title: "HasMinorGridlines (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/HasMinorGridlines%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: VWxMNlqy
---

# HasMinorGridlines (属性)

如果坐标轴有次要网格线，则该值为 **True**。只有主要坐标轴组中的坐标轴才能有网格线。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例判断如果图表工作表 Chart1 数值轴含有次要网格线，则将其颜色设置为绿色。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    if (axis.HasMinorGridlines) {
        axis.MinorGridlines.Border.ColorIndex = 4
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 数值轴是否含有次要网格线。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axis.HasMinorGridlines)
}
```
