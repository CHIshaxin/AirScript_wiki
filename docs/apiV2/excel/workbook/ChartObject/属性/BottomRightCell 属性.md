---
title: "BottomRightCell 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/BottomRightCell%20%E5%B1%9E%E6%80%A7.html
---
# BottomRightCell (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表位于该对象右下角下方的单元格。只读。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中嵌入的第一个图表右下角单元格的地址。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    console.log("The bottom right corner is over cell " + chartobject.BottomRightCell.Address())
}
```

```javascript
/*此示例为第一张工作表中 Chart2 图表右下角的单元格赋值。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects("Chart 2")
    chartobject.BottomRightCell.Value2 = "WPS"
}
```
