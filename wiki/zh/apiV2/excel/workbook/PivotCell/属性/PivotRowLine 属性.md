---
title: "PivotRowLine (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/PivotRowLine%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DXawKETg
---

# PivotRowLine (属性)

在行上返回特定 **[PivotCell](./../../PivotCell/PivotCell 对象.html)** 对象的 PivotLine。只读 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 类型。

## 说明

如果 PivotCell 在行上，则 **PivotRowLine** 返回行的 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 对象。

如果 PivotCell 在列上，则 **PivotRowLine** 返回一个运行时错误。

如果 PivotCell 在数据区域中，则 **PivotRowLine** 返回对应的行的 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 对象。

## 示例

```javascript
/*本示例显示包含活动单元格的数据透视项行轴上数据透视线的PivotLineCells属性中第二个PivotCell应用区域的值。*/
function test() {
    let pvtCell = ActiveCell.PivotCell.PivotRowLine.PivotLineCells.Item(2)
    console.log(pvtCell.Range.Value2)
}
```

```javascript
/*本示例显示包含工作表 Sheet1 中 K10 单元格的数据透视项行轴上数据透视线是否检索所有隐藏的单元格。*/
function test() {
    let pvtLine = Application.Worksheets.Item("Sheet1").Range("K10").PivotCell.PivotRowLine
    console.log(pvtLine.PivotLineCellsFull.Full)
}
```
