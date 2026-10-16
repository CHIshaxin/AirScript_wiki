---
title: "AutoScaling (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/AutoScaling%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DjdkaHKA
---

# AutoScaling (属性)

如果 ET 对三维图表进行缩放，使之与等效的二维图表的大小相近，则该属性值为 **True**。**RightAngleAxes** 属性必须为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例自动对图表工作表Chart1 进行缩放。本示例应在三维图表上运行。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.AutoScaling = true
}
```

```javascript
/*本示例关闭对 Sheet1 工作表第一个内嵌图表的自动缩放。*/
function test() {
    let chart = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart
    chart.AutoScaling = false
}
```
