---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DropLines/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的第一个图表组的垂直线的名称。*/
function test() {
    let droplines = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    console.log(droplines.Name)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个图表的第一个图表组的垂直线的名称。*/
function test() {
    let droplines = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    console.log(`垂直线的名称：${droplines.Name}`)
}
```
