---
title: "BringToFront (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E6%96%B9%E6%B3%95/BringToFront%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BQeuoBQH
---

# BringToFront (方法)

将对象放到 z-次序前面。

## 示例

```javascript
/*此示例将第一张工作表上第一个图表放到 z-次序的前面。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects(1)
    chartobject.BringToFront()
}
```

```javascript
/*此示例将工作表 Sheet1 上图表 Chart 1 放到 z-次序的前面，并将该图表隐藏。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 1")
    chartobject.BringToFront()
    chartobject.Visible = false
}
```
