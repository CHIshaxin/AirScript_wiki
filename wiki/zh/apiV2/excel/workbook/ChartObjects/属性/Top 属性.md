---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B1Lh-aGJ
---

# Top (属性)

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*此示例使工作表 Sheet1 中所有图表的上边界与工作表第五行的上边界对齐。*/
function test() {
    let sheet = Application.Worksheets.Item("Sheet1")
    sheet.ChartObjects().Top = sheet.Rows.Item(5).Top
}
```

```javascript
/*此示例显示图表工作表 Chart1 中图表到该表顶部的距离。*/
function test() {
    let chartobjects = Application.Charts.Item("Chart1").ChartObjects()
    console.log(`到顶部的距离为：${chartobjects.Top} 磅`)
}
```
