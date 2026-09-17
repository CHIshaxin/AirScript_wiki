---
title: "Index 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListRow/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
---
# Index (属性)

返回一个 **Long** 值，它代表 **[ListRow](./../../ListRow/ListRow 对象.html)** 对象在 **[ListRows](./../../ListRows/ListRows 对象.html)** 集合中的索引号。

## 示例

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第二行的索引号。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListRows.Item(2).Index)
}
```

```javascript
/*本示例显示活动工作表的第一个 ListObject 对象各行的索引。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    for (let i = 1; i <= listObj.ListRows.Count; i++) {
        console.log(listObj.ListRows.Item(i).Index)
    }
}
```
