---
title: "TotalsRowRange (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/TotalsRowRange%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bsbub9T5
---

# TotalsRowRange (属性)

返回一个 **Range**，代表指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的“汇总”行（如果有）。只读。

## 示例

```javascript
/*本示例显示活动工作簿的 Sheet1 中默认列表的“汇总”行的地址。该代码可在“汇总”行尚未显示的情况下显示该行。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.ShowTotals = true
    console.log(listObj.TotalsRowRange.Address())
}
```

```javascript
/*本示例显示活动工作表上包含 A3 单元格的列表的汇总行的值。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A3").ListObject
    console.log(listObj.TotalsRowRange.Value2.toString())
}
```
