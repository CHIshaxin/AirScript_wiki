---
title: "ListRow (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListRow/ListRow%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: DBo5kgc4
---

# ListRow (对象)

代表表格中的一行。 **[ListRow](./../ListRow/ListRow 对象.html)** 对象是 **[ListRows](./../ListRows/ListRows 对象.html)** 集合的成员。

## 说明

ListRows 集合包含列表对象中的所有行。**ListRows**

使用 **[ListObject](./../ListObject/ListObject 对象.html)** 对象的 **ListRows** 属性可返回一个 **[ListRows](./../ListRows/ListRows 对象.html)** 集合。

## 示例

```javascript
/*下例给活动工作簿的第一张工作表的默认 ListObject 对象添加新 ListRow 对象。由于未指定位置，因此在表格结束处添加一个新行。*/
function test() {
    let worksheet = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let listRow = worksheet.ListObjects.Item(1).ListRows.Add()
}
```

```javascript
/*本示例将第一张工作表的第一个 ListObject 对象的第三行删除。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    listObj.ListRows.Item(3).Delete()
}
```
