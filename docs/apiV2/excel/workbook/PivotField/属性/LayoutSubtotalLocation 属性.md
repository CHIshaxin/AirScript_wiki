---
title: "LayoutSubtotalLocation 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/LayoutSubtotalLocation%20%E5%B1%9E%E6%80%A7.html
---
# LayoutSubtotalLocation (属性)

返回或设置与指定字段相关（在其上面或下面）的数据透视表字段分类汇总的位置。 **[XlSubtototalLocationType](./../../枚举/XlSubtototalLocationType 枚举.html)** 类型，可读写。

## 说明

可以对任意具有分级显示格式的数据透视表字段设置此属性；但是，只有当指定字段是除最里层（最低级别（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。））行字段之外的其他行字段时，格式才会出现。对于非 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

**LayoutForm**属性确定报表是以表格格式还是以分级显示格式显示。

## 示例

```javascript
/*本示例以分级显示格式显示当前活动工作表上第一个数据透视表中的“name”字段，然后在该字段顶部显示分类汇总。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("name")
    pvtField.LayoutForm = xlOutline
    pvtField.LayoutSubtotalLocation = xlTop
}
```

```javascript
/*本示例显示第一张工作表中数据透视表字段“score”分类汇总的位置是否为“xlAtTop”。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("score")
    console.log(pvtField.LayoutSubtotalLocation == xlAtTop)
}
```
