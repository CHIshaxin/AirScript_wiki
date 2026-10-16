---
title: "ListColumn (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/ListColumn%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: DP3WVgGc
---

# ListColumn (对象)

代表表格中的一列。

## 说明

**[ListColumn](./../ListColumn/ListColumn 对象.html)** 对象是 **[ListColumns](./../ListColumns/ListColumns 对象.html)** 集合的成员。 **[ListColumns](./../ListColumns/ListColumns 对象.html)** 集合包含表格中的所有列（ **[ListObject](./../ListObject/ListObject 对象.html)** 对象）。

使用 **[ListObject](./../ListObject/ListObject 对象.html)** 对象的 ListColumns 属性可返回一个 **[ListColumns](./../ListColumns/ListColumns 对象.html)** 集合。

## 示例

```javascript
/*下例给活动工作簿的第一个工作表的默认 ListObject 对象添加新 ListColumn 对象。由于未指定位置，因此在最右边添加一个新列，并显示列名。*/
function test() {
    let worksheet = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let listColumn = worksheet.ListObjects.Item(1).ListColumns.Add()
    console.log(listColumn.Name)
}
```

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第二列删除。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(2).Delete()
}
```
