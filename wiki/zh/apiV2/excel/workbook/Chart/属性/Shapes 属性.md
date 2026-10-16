---
title: "Shapes (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Shapes%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DZgOMonN
---

# Shapes (属性)

返回一个 **[Shapes](./../../Shapes/Shapes 对象.html)** 集合，它代表图表工作表上所有的形状。只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中Shape对象的数量。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    console.log(chart.Shapes.Count)
}
```

```javascript
/*本示例将工作表 Sheet1 第一个内嵌图表的第一个形状的前景色设置为红色。*/
function test() {
    let chart = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart
    chart.Shapes.Item(1).Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```
