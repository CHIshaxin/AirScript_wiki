---
title: "AxisTitle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/AxisTitle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: v9k_hF69
---

# AxisTitle (属性)

返回一个 **[AxisTitle](./../../AxisTitle/AxisTitle 对象.html)** 对象，该对象表示指定坐标轴的标题。只读。

## 示例

```javascript
/*本示例向图表工作表 Chart1 上图表的分类轴添加坐标轴标签。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasTitle = true
    axis.AxisTitle.Text = "July Sales"
}
```

```javascript
/*本示例向工作表 Sheet1 上第一张图表的数值轴添加坐标轴标签，并将字号设置为 10 磅。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.HasTitle = true
    axis.AxisTitle.Text = "2023年度"
    axis.AxisTitle.Font.Size = 10
}
```
