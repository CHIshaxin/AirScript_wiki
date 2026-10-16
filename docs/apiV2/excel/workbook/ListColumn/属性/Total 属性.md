---
title: "Total 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/%E5%B1%9E%E6%80%A7/Total%20%E5%B1%9E%E6%80%A7.html
---
# Total (属性)

返回 **[ListColumn](./../../ListColumn/ListColumn 对象.html)** 对象的总计行。只读。

## 示例

```javascript
/*本示例显示活动工作表的第一个 ListObject 对象的第二列总计行的值是否大于15。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    console.log(listObj.ListColumns.Item(2).Total.Value2 > 15)
}
```

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第一列的总计行的地址。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListColumns.Item(1).Total.Address())
}
```
