---
title: "Left (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Zckktl1D
---

# Left (属性)

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例将工作表 Sheet1 上第一张图表的标题到图表左边缘的距离设置为标题的宽度。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    charttitle.Left = charttitle.Width
}
```

```javascript
/*本示例显示图表工作表 Chart1 中图表的标题到图表左边缘的距离。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    console.log(`图表工作表 Chart1 中图表的标题到图表左边缘的距离：${charttitle.Left} 磅`)
}
```
