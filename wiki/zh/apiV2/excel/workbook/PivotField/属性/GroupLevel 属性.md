---
title: "GroupLevel (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/GroupLevel%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cuazy9ka
---

# GroupLevel (属性)

返回一组字段中指定字段的位置（如果该字段是分组字段集合中的成员）。只读。

## 说明

该属性对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

处于最高级的父字段（即最左侧的父字段）为第一级，其子字段为第二级，依此类推。

## 示例

```javascript
/*本示例检查包含活动单元格的字段是否为最高一级的父字段，如果是则显示一个消息框。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    if (ActiveCell.PivotField.GroupLevel == 1) {
        console.log("This is the highest-level parent field.")
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的字段“score”的位置。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("score")
    console.log(pvtField.GroupLevel)
}
```
