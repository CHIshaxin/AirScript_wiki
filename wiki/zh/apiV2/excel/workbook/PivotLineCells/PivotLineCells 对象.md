---
title: "PivotLineCells (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLineCells/PivotLineCells%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CWqZ_WtO
---

# PivotLineCells (对象)

特定 PivotLine 的 **[PivotCell](./../PivotCell/PivotCell 对象.html)** 对象的集合。

## 说明

使用 **PivotLineCells** (*index*)方法可以返回或指定集合中特定 **[PivotCell](./../PivotCell/PivotCell 对象.html)** 对象的位置。您也可以指定 **[PivotField](./../PivotField/PivotField 对象.html)** 对象或 PivotField 的名称以返回单个 **[PivotCell](./../PivotCell/PivotCell 对象.html)** 对象。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表列轴上第二条数据透视线的PivotLineCells属性的PivotCell对象的数量。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCells.Count)
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表行轴上第一条数据透视线PivotLineCells属性的第一个PivotCell对象对应的“PivotTable”实体类型是否为“xlPivotCellPivotItem”。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Item(1).PivotCellType == xlPivotCellPivotItem)
}
```
