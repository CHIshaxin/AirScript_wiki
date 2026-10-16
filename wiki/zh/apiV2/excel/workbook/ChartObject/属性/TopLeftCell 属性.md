---
title: "TopLeftCell (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/TopLeftCell%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Mp_uguCZ
---

# TopLeftCell (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表位于指定对象左上角下方的单元格。只读。

## 示例

```javascript
/*此示例显示 Sheet1 上第一张图表左上角下面的单元格的地址。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    console.log("The top left corner is over cell " + chartobject.TopLeftCell.Address())
}
```

```javascript
/*此示例为工作表 Sheet1 上图表 Chart 2 左上角下方的单元格赋值。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    chartobject.TopLeftCell.Value2 = "金山WPS"
}
```
