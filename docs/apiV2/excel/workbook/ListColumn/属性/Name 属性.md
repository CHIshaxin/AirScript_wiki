---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回或设置一个 **String** 值，它代表列表列的名称。

## 说明

该属性值还被用作列表列的显示名称。此名称必须在列表中唯一。

## 示例

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第三列的名称。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    let objListCols = listObj.ListColumns
    console.log(objListCols.Item(3).Name)
}
```

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第四列名称改为“列四”。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(4).Name = "列四"
}
```
