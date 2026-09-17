---
title: "ListColumns 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumns/ListColumns%20%E5%AF%B9%E8%B1%A1.html
---
# ListColumns (对象)

指定的 **[ListObject](./../ListObject/ListObject 对象.html)** 对象中所有 **[ListColumn](./../ListColumn/ListColumn 对象.html)** 对象的集合。

## 说明

每个 **[ListColumn](./../ListColumn/ListColumn 对象.html)** 对象都代表表格中的一列。

该列的名称会自动生成。在添加完该列后可更改其名称。

使用 ListObject 对象的 **ListColumns** 属性可返回 **[ListColumns](./../ListColumns/ListColumns 对象.html)** 集合。

## 示例

```javascript
/*下例给工作簿的第一张工作表的默认 ListObject 对象添加一个新列。由于未指定位置，因此在最右边添加一个新列。*/
function test() {
    let newColumn = Application.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Add()
}
```

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第四列删除。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(4).Delete()
}
```
