---
title: "Function (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/Function%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DU4npi8G
---

# Function (属性)

返回或设置对数据透视表字段汇总时所使用的函数（仅用于数据字段）。 **[XlConsolidationFunction](./../../枚举/XlConsolidationFunction 枚举.html)** 类型，可读写。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，本属性是只读的，并且总是返回 **xlUnknown**。对于其他的数据源，本属性不能被设置为 **xlUnknown**。

## 示例

```javascript
/*本示例将工作表 Sheet1 中数据透视表的字段“最大值项:grades”汇总时使用的函数设置为“xlSum”。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("最大值项:grades")
    pvtField.Function = xlSum
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“平均值项:id”汇总时使用的函数是否为“xlAverage”。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("平均值项:id").Function == xlAverage)
}
```
