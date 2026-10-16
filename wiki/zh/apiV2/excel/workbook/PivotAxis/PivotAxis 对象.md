---
title: "PivotAxis (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotAxis/PivotAxis%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Dr3fFs5Z
---

# PivotAxis (对象)

**[PivotAxis](./../PivotAxis/PivotAxis 对象.html)** 对象用于在数据透视表中进行不对称深化。

## 说明

**[PivotAxis](./../PivotAxis/PivotAxis 对象.html)** 对象包含诸如 **PivotRowAxis** 和 **PivotRowAxis**之类的属性，用于处理数据透视表中的行和列。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第二条数据透视线的PivotLineCells属性中第一个PivotCell的数据透视表项的值。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCells.Item(1).PivotItem.Value)
}
```

```javascript
/*本示例显示活动工作表上第一张数据透视表中列轴上数据透视线的数量。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotColumnAxis.PivotLines.Count)
}
```
