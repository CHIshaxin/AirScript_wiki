---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

在列表中删除数据列。

## 说明

此方法不从工作表中删除列。如果列表链接到 SharePoint Foundation 网站，则无法从服务器中删除该列，并产生一个错误。

## 示例

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第一列删除。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(1).Delete()
}
```

```javascript
/*本示例将第一张工作表的第一个 ListObject 对象的第三列删除。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    listObj.ListColumns.Item(3).Delete()
}
```
