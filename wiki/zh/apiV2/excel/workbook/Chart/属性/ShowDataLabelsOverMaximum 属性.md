---
title: "ShowDataLabelsOverMaximum (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ShowDataLabelsOverMaximum%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: PqYXuKm2
---

# ShowDataLabelsOverMaximum (属性)

返回或设置一个布尔值，该值表示在数值大于数值轴上的最大值时是否显示数据标签。可读/写 **Boolean** 类型。

## 说明

如果您更改数值轴，使其变得比数据点的大小还小，则可以使用此属性设置是否显示数据标签。此属性只适用于二维图表。

## 示例

```javascript
/*本示例将图表工作表 Chart1 设置为当值大于数值轴上的最大值时显示数据标签。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.ShowDataLabelsOverMaximum = true
}
```

```javascript
/*本示例显示工作表 Sheet2 中第一个内嵌图表当值大于数值轴上的最大值时是否显示数据标签。*/
function test() {
    let chart = Application.Sheets.Item("Sheet2").ChartObjects(1).Chart
    if (chart.ShowDataLabelsOverMaximum) {
        console.log("图表当值大于数值轴上的最大值时显示数据标签")
    } else {
        console.log("图表当值大于数值轴上的最大值时不显示数据标签")
    }
}
```
