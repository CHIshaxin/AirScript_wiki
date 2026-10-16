---
title: "Orientation 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
---
# Orientation (属性)

返回或设置一个 **Variant** 值，它代表文本方向。

## 说明

此属性的值可以设置为 -90 到 90 度的整数值，或设置为 **[XlOrientation](./../../枚举/XlOrientation 枚举.html)** 常量之一。

## 示例

```javascript
/*本示例判断如果图表工作表 Chart1 的标题可见，则将标题的方向设置为向上排列。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    if (chart.HasTitle) {
        chart.ChartTitle.Orientation = xlUpward
    }
}
```

```javascript
/*本示例将工作表 Sheet1 中图表标题的方向设置为 -90 度。*/
function test() {
    let chtitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    chtitle.Orientation = -90
}
```
