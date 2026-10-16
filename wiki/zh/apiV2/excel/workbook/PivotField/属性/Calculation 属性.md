---
title: "Calculation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/Calculation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: ChSep-QG
---

# Calculation (属性)

返回或设置一个 **[XlPivotFieldCalculation](./../../枚举/XlPivotFieldCalculation 枚举.html)** 值，它代表指定字段执行的计算类型。此属性仅对数据字段有效。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性仅能返回或被设置为 **xlNormal**。

## 示例

```javascript
/*本示例将第一张工作表上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“name”字段，并将基准数据项设为“张五”。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("求和项:score")
    pvtField.Calculation = xlDifferenceFrom
    pvtField.BaseField = "name"
    pvtField.BaseItem = "张五"
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“求和项:期中成绩”的计算类型是否为“xlPercentOf”。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("求和项:期中成绩").Calculation == xlPercentOf)
}
```
