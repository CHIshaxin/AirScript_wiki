---
title: "IsCalculated 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/IsCalculated%20%E5%B1%9E%E6%80%A7.html
---
# IsCalculated (属性)

如果数据透视表字段或数据透视表项为计算字段或计算项，则此属性为 **True**。**Boolean** 类型，只读。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终返回 **False**。

## 示例

```javascript
/*本示例显示活动工作表中数据透视表的字段“name”的第一个数据项是否为计算项。*/
function test() {
    let pvtItem = ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotItems(1)
    console.log(pvtItem.IsCalculated)
}
```

```javascript
/*本示例判断工作表 Sheet1 上第一张数据透视表的字段“name”中名为“公式1”的数据项是否为计算项，并通知用户。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems("公式1")
    if (pvtItem.IsCalculated) {
        console.log("数据项是计算项")
    } else {
        console.log("数据项不是计算项")
    }
}
```
