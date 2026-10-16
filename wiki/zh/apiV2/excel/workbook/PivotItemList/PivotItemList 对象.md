---
title: "PivotItemList (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItemList/PivotItemList%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BG8AWta0
---

# PivotItemList (对象)

指定的数据透视表中所有 **[PivotItem](./../PivotItem/PivotItem 对象.html)** 对象的集合。

## 说明

每个 **PivotItem** 代表数据透视表字段中的一个项。

使用 **[PivotCell](./../PivotCell/PivotCell 对象.html)** 对象的 **RowItems** 或 **ColumnItems** 属性可返回一个 **[PivotItemList](./../PivotItemList/PivotItemList 对象.html)** 集合。

一旦返回了 **[PivotItemList](./../PivotItemList/PivotItemList 对象.html)** 集合，您就可以使用 **Item** 方法来标识某个特定的 **PivotItem** 列表。

## 示例

```javascript
/*本示例显示包含单元格 L10 的数据透视项行轴上第一个数据项的标签文本。*/
function test() {
    // Identify contents associated with PivotItemList.
    console.log("Contents associated with cell L10: " + Application.Range("L10").PivotCell.RowItems.Item(1).Caption)
}
```

```javascript
/*本示例将工作表 Sheet1 上包含单元格 L4 的数据透视项行轴上第一个数据项的详细信息设置为不可见。*/
function test() {
    let pvtItem = Application.Worksheets.Item("Sheet1").Range("L4").PivotCell.RowItems.Item(1)
    pvtItem.ShowDetail = false
}
```
