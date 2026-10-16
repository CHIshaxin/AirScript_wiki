---
title: "PivotLineCellsFull 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLine/%E5%B1%9E%E6%80%A7/PivotLineCellsFull%20%E5%B1%9E%E6%80%A7.html
---
# PivotLineCellsFull (属性)

检索所有数据透视线单元格，包括以压缩形式隐藏的单元格。 此为只读属性。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第一条数据透视线检索单元格的数量。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCellsFull.Count)
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表中行轴上第二条数据透视线是否检索所有隐藏的单元格。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCellsFull.Full)
}
```
