---
title: "DataField (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/DataField%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: OqcVcgPl
---

# DataField (属性)

返回一个 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象，该对象与所选数据字段相对应。

## 说明

如果 **[PivotCell](./../../PivotCell/PivotCell 对象.html)** 对象不是 **[XlPivotCellType](./../../枚举/XlPivotCellType 枚举.html)** 允许的常量：**xlPivotCellTypeDataField**、**xlPivotCellTypeSubtotal** 或 **xlPivotCellTypeGrandTotal** 之一，那么该属性将返回一个错误。

## 示例

```javascript
/*本示例确定单元格 I2 是否在数据透视表的数据字段中，然后通过通知用户来返回对应于该单元格的数据透视表字段的值，或处理运行错误。本示例假定数据透视表位于活动工作表上。*/
function test() {
    try {
        console.log(Application.Range("I2").PivotCell.DataField.Value)
    } catch (exception) {
        console.log("The selected range is not in the data field of the PivotTable.")
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表行轴上第一条数据透视线PivotLineCells属性的第一个PivotCell对象对应字段汇总时所使用的函数是否为 xlSum。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Item(1).DataField.Function == xlSum)
}
```
