---
title: "ShowTableStyleLastColumn 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ShowTableStyleLastColumn%20%E5%B1%9E%E6%80%A7.html
---
# ShowTableStyleLastColumn (属性)

返回或设置是否显示指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的最后一列。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例显示活动工作表上包含 B2 单元格的列表是否展示最后一列。*/
function test() {
    let listObj = Application.ActiveSheet.Range("B2").ListObject
    console.log(listObj.ShowTableStyleLastColumn)
}
```

```javascript
/*本示例将工作表 Sheet1 上第一张列表设置为展示最后一列。*/
function test() {
    let listObj = Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.ShowTableStyleLastColumn = true
}
```
