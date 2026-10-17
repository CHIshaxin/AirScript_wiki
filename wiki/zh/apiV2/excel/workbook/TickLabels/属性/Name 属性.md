---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: EjdVHRyN
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*下例显示 Sheet1 上的第一个图表分类轴刻度线标签的名称。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    console.log(tickLabels.Name)
}
```

```javascript
/*以下示例显示 Chart1 上数值轴刻度线标签的名称。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).TickLabels
    console.log(tickLabels.Name)
}
```
