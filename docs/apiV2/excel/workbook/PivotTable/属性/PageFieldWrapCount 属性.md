---
title: "PageFieldWrapCount 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/PageFieldWrapCount%20%E5%B1%9E%E6%80%A7.html
---
# PageFieldWrapCount (属性)

返回或设置数据透视表中每行或每列的页字段数目。**Long** 类型，可读写。

## 示例

```javascript
/*本示例使数据透视表先在一行中显示三个页字段，然后再显示下一个新行。*/
function test() {
    let pTab = Application.Worksheets.Item(1).PivotTables("Pivot1")
    pTab.PageFieldOrder = xlOverThenDown
    pTab.PageFieldWrapCount = 3
}
```

```javascript
/*本示例将数据透视表的 PageFieldWrapCount 属性设置为 5，并弹窗显示该属性的值*/
function test() {
    ActiveSheet.PivotTables(1).PageFieldWrapCount = 5
    console.log(ActiveSheet.PivotTables(1).PageFieldWrapCount)
}
```
