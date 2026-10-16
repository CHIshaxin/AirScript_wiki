---
title: "Full (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLineCells/%E5%B1%9E%E6%80%A7/Full%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DJTh7Oqh
---

# Full (属性)

如果为 **True**，则会检索数据透视线中的所有 **PivotLineCells** (包括以压缩形式) 隐藏的那些单元格。 **Boolean** 类型，只读。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第一条数据透视线是否检索所有PivotLineCells隐藏的单元格。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Full)
}
```

```javascript
/*本示例显示工作表 Sheet1 上数据透视表中行轴上第二条数据透视线是否检索所有PivotLineCells隐藏的单元格。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCells.Full)
}
```
