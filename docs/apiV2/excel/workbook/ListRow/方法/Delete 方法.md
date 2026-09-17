---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListRow/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除列表行的单元格并向上移动被删除行下面的任何剩余单元格。即使列表被链接到 SharePoint 网站，您也可以删除该列表中的行。但是，在同步更改之前，SharePoint 网站上的列表将不会进行更新。

## 示例

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第一行删除。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.ListRows.Item(1).Delete()
}
```

```javascript
/*本示例将第一张工作表的第一个 ListObject 对象的第二行删除。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    listObj.ListRows.Item(2).Delete()
}
```
