---
title: "PivotLineCells 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLine/%E5%B1%9E%E6%80%A7/PivotLineCells%20%E5%B1%9E%E6%80%A7.html
---
# PivotLineCells (属性)

返回 PivotLine 中 **[PivotCell](./../../PivotCell/PivotCell 对象.html)** 对象的集合。只读。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表中行轴上第一条数据透视线“PivotCell”对象的数量。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Count)
}
```

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第一条数据透视线的PivotLineCells属性中第一个PivotCell应用区域的值。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Item(1).Range.Value2)
}
```
