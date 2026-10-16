---
title: "ListObject (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/ListObject%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B2Zm-2NU
---

# ListObject (对象)

代表工作表中的表格。

## 说明

**[ListObject](./../ListObject/ListObject 对象.html)** 对象是 **[ListObjects](./../ListObjects/ListObjects 对象.html)** 集合的成员。 **[ListObjects](./../ListObjects/ListObjects 对象.html)** 集合包含工作表上所有的列表对象。

使用 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象的 ListObjects 属性可返回一个 **[ListObjects](./../ListObjects/ListObjects 对象.html)** 集合。

## 示例

```javascript
/*本示例给活动工作簿的第一张工作表的默认 ListObject 对象添加新的 ListRow 对象。*/
function test() {
    let sheet = Application.ActiveWorkbook.Worksheets.Item(1)
    sheet.ListObjects.Item(1).ListRows.Add()
}
```

```javascript
/*本示例为工作表 Sheet1 上第一张列表添加新列，并设置该列表的表样式。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.ListColumns.Add()
    listObj.TableStyle = "TableStylePreset5_Accent1"
}
```
