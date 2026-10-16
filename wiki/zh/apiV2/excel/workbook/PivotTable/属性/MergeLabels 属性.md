---
title: "MergeLabels (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/MergeLabels%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BSyp2juF
---

# MergeLabels (属性)

如果指定的数据透视表的外部行项、列项、分类汇总和总计标志使用合并单元格，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例使第一张工作表上的第一个数据透视表使用合并单元格外部行项、列项、分类汇总和总计的标志。*/
function test() {
    Application.Worksheets.Item(1).PivotTables(1).MergeLabels = true
}
```

```javascript
/*本示例显示数据透视表 MergeLabels 属性的数据类型，并设置为 false。*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).MergeLabels)
    ActiveSheet.PivotTables(1).MergeLabels = false
}
```
