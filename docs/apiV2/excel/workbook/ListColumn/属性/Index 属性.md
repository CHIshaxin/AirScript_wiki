---
title: "Index 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
---
# Index (属性)

返回一个 **Long** 值，它代表 **[ListColumn](./../../ListColumn/ListColumn 对象.html)** 对象在 **[ListColumns](./../../ListColumns/ListColumns 对象.html)** 集合中的索引号。

## 示例

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第二列的索引号。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListColumns.Item(2).Index)
}
```

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象的第三列索引。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListColumns.Item(3).Index)
}
```
