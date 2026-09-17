---
title: "HasTitle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/HasTitle%20%E5%B1%9E%E6%80%A7.html
---
# HasTitle (属性)

如果坐标轴或图表有可见标题，则为 **True**。**Boolean** 类型，可读写。

## 说明

图表标题由 **[ChartTitle](./../../ChartTitle/ChartTitle 对象.html)** 对象代表。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 的标题，并将标题的字体大小设为20。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.HasTitle = true
    chart.ChartTitle.Font.Size = 20
}
```

```javascript
/*本示例设置第一张工作表第一个嵌入式图表的标题为不可见。*/
function test() {
    let chart = Application.Sheets.Item(1).ChartObjects(1).Chart
    chart.HasTitle = false
}
```
