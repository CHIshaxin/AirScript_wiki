---
title: "ListObjects (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObjects/ListObjects%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: DzEHVwPV
---

# ListObjects (对象)

工作表上所有 **[ListObject](./../ListObject/ListObject 对象.html)** 对象的集合。每个 **[ListObject](./../ListObject/ListObject 对象.html)** 对象都代表工作表中的一个表格。

## 说明

使用 Worksheet 对象的 **[ListObjects](./../ListObjects/ListObjects 对象.html)** 属性可返回 **[ListObjects](./../ListObjects/ListObjects 对象.html)** 集合。

## 示例

```javascript
/*本示例创建一个新 ListObjects 集合，该集合代表第一张工作表中所有的表格，并显示表格数量。*/
function test() {
    let listObjects = Application.Worksheets.Item(1).ListObjects
    console.log(listObjects.Count)
}
```

```javascript
/*本示例显示活动工作表上第二个列表的 SourceType 属性是否为 xlSrcRange。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(2)
    console.log(listObj.SourceType == xlSrcRange)
}
```
