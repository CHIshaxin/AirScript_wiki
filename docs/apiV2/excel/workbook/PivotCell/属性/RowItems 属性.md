---
title: "RowItems 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/RowItems%20%E5%B1%9E%E6%80%A7.html
---
# RowItems (属性)

返回一个 **[PivotItemList](./../../PivotItemList/PivotItemList 对象.html)** 集合，该集合对应于分类轴上表示选定单元格的项目。

## 示例

```javascript
/*本示例判断单元格 L9 中的数据项是否在第一个行字段中的“name”项之下，并通知用户。本示例假定数据透视表位于活动工作表上，并且工作表的列 L 包含数据透视表的行项目。*/
function test() {
    // Determine if there is a match between the item and column field.
    if (Application.Range("L9").PivotCell.RowItems.Item(1).Parent.Name == "name") {
        console.log("Cell L9 is a member of the 'name' row field.")
    } else {
        console.log("Cell L9 is not a member of the 'name' row field.")
    }
}
```

```javascript
/*本示例显示包含活动单元格的数据透视项行轴上第二个数据项的值。*/
function test() {
    console.log(ActiveCell.PivotCell.RowItems.Item(2).Value)
}
```
