---
title: "ListColumns 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ListColumns%20%E5%B1%9E%E6%80%A7.html
---
# ListColumns (属性)

返回一个 **[ListColumns](./../../ListColumns/ListColumns 对象.html)** 集合，该集合代表 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象中的所有列。只读。

## 示例

```javascript
/*本示例显示了通过对 ListColumns 属性的调用所创建的 ListColumns 集合对象中第二列的名称。要让本示例运行，Sheet1 工作表必须包含一个表格。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    let objListCols = listObj.ListColumns
    console.log(objListCols.Item(2).Name)
}
```

```javascript
/*本示例显示活动工作表上包含 A3 单元格的列表中列的数量。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A3").ListObject
    console.log(listObj.ListColumns.Count)
}
```
