---
title: "IsCalculated (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/IsCalculated%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DP_fOt-I
---

# IsCalculated (属性)

如果数据透视表字段为计算字段或计算项，则此属性为 **True**。**Boolean** 类型，只读。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终返回 **False**。

## 示例

```javascript
/*本示例判断如果数据透视表包含任何计算字段，此示例将禁用“数据透视表字段”对话框。*/
function test() {
    let pt = Worksheets.Item(1).PivotTables("数据透视表1")
    for (let i = 1; i <= pt.PivotFields().Count; i++) {
        if (pt.PivotFields(i).IsCalculated) {
            pt.EnableFieldDialog = false
        }
    }
}
```

```javascript
/*本示例显示活动工作表中数据透视表的字段“期中成绩”是否为计算字段。*/
function test() {
    console.log(ActiveSheet.Range("I1").PivotTable.PivotFields("期中成绩").IsCalculated)
}
```
