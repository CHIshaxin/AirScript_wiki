---
title: "ListRows (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ListRows%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CqXNPl9Y
---

# ListRows (属性)

返回一个 **[ListRows](./../../ListRows/ListRows 对象.html)** 对象，该对象代表 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象中的所有数据行。只读。

## 说明

返回的 **[ListRows](./../../ListRows/ListRows 对象.html)** 对象不包括标题行、总计行或插入行。

## 示例

```javascript
/*本示例删除了由活动工作簿上工作表 Sheet1 中第一张列表的 ListRows 集合中的 iRowNumber 指定的行。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    let objListRows = listObj.ListRows
    let iRowNumber = 1
    if ((iRowNumber !== 0) && (iRowNumber < objListRows.Count - 1)) {
        objListRows.Item(iRowNumber).Delete()
    }
}
```

```javascript
/*本示例为活动工作表上包含 B3 单元格的列表添加数据行。*/
function test() {
    let listObj = Application.ActiveSheet.Range("B3").ListObject
    listObj.ListRows.Add()
}
```
