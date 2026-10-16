---
title: "PivotItem (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/PivotItem%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BybCmgn1
---

# PivotItem (属性)

返回一个 **[PivotItem](./../../PivotItem/PivotItem 对象.html)** 对象，它代表指定区域左上角所在的数据透视表项。

## 示例

```javascript
/*此示例显示包含 Sheet1 中活动单元格的数据透视表数据项的名称。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    console.log ("The active cell is in the item " + ActiveCell.PivotItem.Name)
}
```

```javascript
/*本示例将 A3 单元格所对应的数据项名称设置为“表一”，并显示*/
function test() {
    Range("A3").PivotItem.Value = "表一"
    console.log(Range("A3").PivotItem.Value)
}
```
