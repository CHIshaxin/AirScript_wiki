---
title: "ColumnItems 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/ColumnItems%20%E5%B1%9E%E6%80%A7.html
---
# ColumnItems (属性)

返回一个 **[PivotItemList](./../../PivotItemList/PivotItemList 对象.html)** 集合，该集合对应于表示选定区域的列轴上的项目。

## 示例

```javascript
/*本示例判断单元格 J5 中的数据项是否在列字段“name”之下，并通知用户。本示例假定数据透视表位于活动工作表上，并且第 J 列包含数据透视表的列字段。*/
function test() {
    // Determine if there is a match between the item and column field.
    if (Application.Range("J5").PivotCell.ColumnItems.Item(1).Parent.Name == "name") {
        console.log("Item in J5 is a member of the 'name' column field.")
    } else {
        console.log("Item in J5 is not a member of the 'name' column field.")
    }
}
```

```javascript
/*本示例显示包含活动单元格的数据透视表项列轴上第二个数据项的位置。*/
function test() {
    console.log(ActiveCell.PivotCell.ColumnItems.Item(2).Position)
}
```
