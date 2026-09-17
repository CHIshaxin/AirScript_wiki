---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumns/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第二列的名称。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    let objListCols = listObj.ListColumns
    console.log(objListCols.Item(2).Name)
}
```

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第三列删除。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(3).Delete()
}
```
