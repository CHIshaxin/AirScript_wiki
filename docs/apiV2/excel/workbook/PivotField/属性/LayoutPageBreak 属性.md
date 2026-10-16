---
title: "LayoutPageBreak 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/LayoutPageBreak%20%E5%B1%9E%E6%80%A7.html
---
# LayoutPageBreak (属性)

如果在字段的每项后都插入了分页符，则该值为 **True**。 默认值为 **False**。**Boolean** 类型，可读写。

## 说明

虽然可以对任意数据透视表字段设置此属性，但只有当指定字段是除最里层（最低级别（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。））行字段之外的其他行字段时，打印选项才会出现。对于非 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

## 示例

```javascript
/*本示例在活动工作表的第一个数据透视表的“name”字段之后添加一个分页符。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("name").LayoutPageBreak = true
}
```

```javascript
/*本示例显示第一张工作表中数据透视表的字段“score”之后是否添加分页符。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("score")
    console.log(pvtField.LayoutPageBreak)
}
```
