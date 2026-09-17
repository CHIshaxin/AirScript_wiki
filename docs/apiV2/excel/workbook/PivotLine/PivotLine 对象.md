---
title: "PivotLine 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLine/PivotLine%20%E5%AF%B9%E8%B1%A1.html
---
# PivotLine (对象)

**[PivotLine](./../PivotLine/PivotLine 对象.html)** 对象是 ET 数据透视表中的行或列的线条。

## 说明

PivotLine 只包含可见项，因此 **[PivotLine](./../PivotLine/PivotLine 对象.html)** 集合中不存在折叠的项目子项以及隐藏级别中的项目。

PivotLine 在所有位置始终具有一个 PivotItem。这意味着与普通 PivotLine 相比，代表数据透视表中分类汇总的 PivotLine 包含较少的 PivotItem。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表中列轴上第二条数据透视线的PivotLineCells属性中PivotCell对象的数量。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotColumnAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCells.Count)
}
```

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第九条数据透视线的类型是否为“xlPivotLineGrandTotal”。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(9)
    console.log(pvtLine.LineType == xlPivotLineGrandTotal)
}
```
