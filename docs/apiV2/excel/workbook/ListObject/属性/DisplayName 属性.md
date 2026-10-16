---
title: "DisplayName 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/DisplayName%20%E5%B1%9E%E6%80%A7.html
---
# DisplayName (属性)

返回或设置指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的显示名称。可读/写 **String** 类型。

## 示例

```javascript
/*本示例显示活动工作表上包含 A2 单元格的列表的展示名称。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A2").ListObject
    console.log(listObj.DisplayName)
}
```

```javascript
/*本示例设置工作表 Sheet1 上第一张列表的展示名称。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.DisplayName = "ListObject_1"
}
```
