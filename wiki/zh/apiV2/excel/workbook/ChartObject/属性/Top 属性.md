---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CImUv-VI
---

# Top (属性)

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*此示例显示工作表 Sheet1 上图表 Chart 2 到工作表第一行顶部的距离。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    console.log(`图表 Chart 2 到工作表第一行顶部的距离：${chartobject.Top} 磅`)
}
```

```javascript
/*此示例使工作表 Sheet1 中第一张图表的上边界与第三行的上边界对齐。*/
function test() {
    let sheet = Application.Worksheets.Item("Sheet1")
    sheet.ChartObjects(1).Top = sheet.Rows.Item(3).Top
}
```
