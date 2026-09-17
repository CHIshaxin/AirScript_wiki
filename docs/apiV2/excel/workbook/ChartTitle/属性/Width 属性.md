---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

返回对象的宽度（以磅为单位）。只读。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个图表的标题宽度。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    console.log(`图表标题的宽度：${charttitle.Width} 磅`)
}
```

```javascript
/*此示例显示图表工作表 Chart1 中图表标题的宽度是否为 62 磅。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    console.log(charttitle.Width == 62)
}
```
