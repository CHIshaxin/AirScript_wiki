---
title: "PivotField (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/PivotField%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CW84LAyV
---

# PivotField (属性)

返回一个 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象，它代表指定区域左上角所在的数据透视表字段。

## 示例

```javascript
/*此示例显示包含活动单元格的数据透视项的数据透视表字段的名称。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    console.log("The active cell is in the field " + ActiveCell.PivotCell.PivotField.Name)
}
```

```javascript
/*本示例将工作表 Sheet1 上第一张数据透视表行轴上第一条数据透视线PivotLineCells属性的第一个PivotCell对应透视表字段的区域选中。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(1)
    pvtLine.PivotLineCells.Item(1).PivotField.DataRange.Select()
}
```
