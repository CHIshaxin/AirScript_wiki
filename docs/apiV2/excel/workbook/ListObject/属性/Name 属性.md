---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回或设置一个 **String** 值，它代表 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的名称。

## 说明

此名称只用作 **[ListObjects](./../../ListObjects/ListObjects 对象.html)** 集合对象的 **Item**属性的唯一标识符。只能通过对象模型设置此属性。

默认情况下，每个 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象名称都以“List”开头，后面接着数字（无空格）。如果试图将 **Name** 属性设置为其他 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象已经使用的名称，将产生运行时错误。

## 示例

```javascript
/*本示例显示活动工作簿的 Sheet1 中的默认 ListObject 对象的名称。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.Name)
}
```

```javascript
/*本示例将活动工作表上包含 A1 单元格的列表的名称设置为“ListObj_1”。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A1").ListObject
    listObj.Name = "ListObj_1"
}
```
