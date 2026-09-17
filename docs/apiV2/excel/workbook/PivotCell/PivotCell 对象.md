---
title: "PivotCell 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/PivotCell%20%E5%AF%B9%E8%B1%A1.html
---
# PivotCell (对象)

代表数据透视表中的一个单元格。

## 说明

使用 **[Range](./../Range/Range 对象.html)** 集合的 **PivotCell** 属性可返回一个 **[PivotCell](./../PivotCell/PivotCell 对象.html)** 对象。

返回 **[PivotCell](./../PivotCell/PivotCell 对象.html)** 对象后，可以使用 **ColumnItems** 或 **RowItems** 属性来确定 **[PivotItems](./../PivotItems/PivotItems 对象.html)** 集合，对应于代表所选编号的列轴或行轴上的项目。

返回了 **[PivotCell](./../PivotCell/PivotCell 对象.html)** 对象后，可以使用 **PivotCellType** 属性来确定某个特定区域是什么单元格类型。

## 示例

```javascript
/*本示例确定数据透视表中的单元格 J6 是否是一个数据项，并通知用户。本示例假定数据透视表位于活动工作表上。如果单元格 J6 不在数据透视表中，则本示例处理运行错误。*/
function test() {
    try {
        // Determine if cell J6 is a data item in the PivotTable.
        if (Application.Range("J6").PivotCell.PivotCellType == xlPivotCellValue) {
            console.log("The cell at J6 is a data item.")
        } else {
            console.log("The cell at J6 is not a data item.")
        }
    }
    catch (exception) {
        console.log("The chosen cell is not in a PivotTable.")
    }
}
```

```javascript
/*此示例确定单元格 L13 的数据项所在的列字段。然后判断列字段标题是否与“name”相匹配，并通知用户。此示例假定数据透视表位于活动工作表上，并且工作表的 L 列包含数据透视表的列字段。*/
function test() {
    // Determine if there is a match between the item and column field.
    if (Application.Range("L13").PivotCell.ColumnItems.Item(1).Parent.Name == "name") {
        console.log("Item in L13 is a member of the 'name' column field.")
    } else {
        console.log("Item in L13 is not a member of the 'name' column field.")
    }
}
```
