---
title: "BaseItem 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/BaseItem%20%E5%B1%9E%E6%80%A7.html
---
# BaseItem (属性)

返回或设置自定义计算基准字段的数据项，仅对数据字段有效。**Variant** 类型，可读写。

## 说明

该属性对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

## 示例

```javascript
/*本示例将第一张工作表上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“name”字段，并将基准数据项设为“张三”。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("求和项:score")
    pvtField.Calculation = xlDifferenceFrom
    pvtField.BaseField = "name"
    pvtField.BaseItem = "张三"
}
```

```javascript
/*本示例显示活动工作表上数据透视表中的数据字段“求和项:期中成绩”的基准数据项的名称。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("求和项:期中成绩")
    console.log(pvtField.BaseItem.Name)
}
```
