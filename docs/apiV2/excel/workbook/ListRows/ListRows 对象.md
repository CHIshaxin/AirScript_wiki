---
title: "ListRows 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListRows/ListRows%20%E5%AF%B9%E8%B1%A1.html
---
# ListRows (对象)

指定的 **[ListObject](./../ListObject/ListObject 对象.html)** 对象中 **[ListRow](./../ListRow/ListRow 对象.html)** 对象的集合。

## 说明

每一个 **[ListRow](./../ListRow/ListRow 对象.html)** 对象都代表表格中的一行。

使用 **[ListObject](./../ListObject/ListObject 对象.html)** 对象的 **ListRows** 属性可返回 **[ListRows](./../ListRows/ListRows 对象.html)** 集合。

## 示例

```javascript
/*下例给第一张工作表的默认 ListObject 对象添加新行。由于未指定位置，因此在表格结束处添加一个新行。*/
function test() {
    let row = Application.Worksheets.Item(1).ListObjects.Item(1).ListRows.Add()
}
```

```javascript
/*本示例显示活动工作表的第一个 ListObject 对象中行的数量。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(1)
    console.log(listObj.ListRows.Count)
}
```
