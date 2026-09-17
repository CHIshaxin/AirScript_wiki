---
title: "PivotLines 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLines/PivotLines%20%E5%AF%B9%E8%B1%A1.html
---
# PivotLines (对象)

**[PivotLines](./../PivotLines/PivotLines 对象.html)** 对象是数据透视表中线条的集合，其中包含数据透视表中行或列上的所有线条。每个线条都是一组 PivotCells。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中行轴上第一条数据透视线的PivotLineCells属性第二个PivotCell对应数据透视表项的名称。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines.Item(1)
    console.log(pvtLine.PivotLineCells.Item(2).PivotItem.Name)
}
```

```javascript
/*本示例显示活动工作表上数据透视表中列轴上数据透视线数量。*/
function test() {
    let pvtLines = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines
    console.log(`列轴上数据透视线的数量:${pvtLines.Count}`)
}
```
