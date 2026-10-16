---
title: "PivotCellType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/PivotCellType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D3TlIeD4
---

# PivotCellType (属性)

返回一个 **[XlPivotCellType](./../../枚举/XlPivotCellType 枚举.html)** 常量，该常量标识单元格所对应的数据透视表实体。只读。

## 说明

| **[XlPivotCellType](./../../枚举/XlPivotCellType 枚举.html)** 可为以下 **[XlPivotCellType](./../../枚举/XlPivotCellType 枚举.html)** 常量之一。 |
| --- |
| **xlPivotCellBlankCell**数据透视表中的一个结构空白单元格。 |
| **xlPivotCellCustomSubtotal** 自定义分类汇总的行或列区域中的一个单元格。 |
| **xlPivotCellDataField**一个数据字段标签（不是“数据”按钮）。 |
| **xlPivotCellDataPivotField**“数据”按钮。 |
| **xlPivotCellGrandTotal**总计行或列区域中的一个单元格。 |
| **xlPivotCellPageFieldItem**显示页字段的选定项的单元格。 |
| **xlPivotCellPivotField**字段的按钮（不是“数据”按钮）。 |
| **xlPivotCellPivotItem**不是分类汇总、总计、自定义分类汇总或空行的行或列区域中的单元格。 |
| **xlPivotCellSubtotal**分类汇总的行或列区域中的单元格。 |
| **xlPivotCellValue**数据区区域（除空行以外）中的任一单元格。 |

## 示例

```javascript
/*本示例确定数据透视表中的单元格 I5 是否是一个数据项，并通知用户。本示例假定数据透视表位于活动工作表上。如果单元格 I5 不在数据透视表中，则本示例处理运行错误。*/
function test() {
    try {
        // Determine if cell I5 is a data item in the PivotTable.
        if (Application.Range("I5").PivotCell.PivotCellType == xlPivotCellValue) {
            console.log("The cell at I5 is a data item.")
        } else {
            console.log("The cell at I5 is not a data item.")
        }
    }
    catch (exception) {
        console.log("The chosen cell is not in a PivotTable.")
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表行轴上第一条数据透视线PivotLineCells属性的第一个PivotCell类型是否为“xlPivotCellPivotItem”。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Item(1).PivotCellType == xlPivotCellPivotItem)
}
```
