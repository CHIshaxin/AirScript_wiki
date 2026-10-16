---
title: "Position (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/Position%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BO7unW2k
---

# Position (属性)

返回或设置一个 **Long** 值，它代表当前正在显示数据项时，该数据项字段中数据项的位置。

## 示例

```javascript
/*此示例显示包含活动单元格的数据透视表项的位置数字。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    console.log("The active item is in position number " + ActiveCell.PivotItem.Position)
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张数据透视表的字段“name”的数据项“张三”的位置设置为第一个。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems("张三")
    pvtItem.Position = 1
}
```
