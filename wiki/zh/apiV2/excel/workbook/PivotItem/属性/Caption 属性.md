---
title: "Caption (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/Caption%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DgA8YS2X
---

# Caption (属性)

返回一个 **String** 值，它代表数据透视项的标签文本。

## 说明

下表显示 **Caption** 属性及其相关属性的示例值，假设存在一个具有唯一名称为“[Europe].[France].[Paris]”的 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在一个包含名为“Paris”项的非 OLAP 数据源。

| 属性 | 值（OLAP 数据源） | 值（非 OLAP 数据源） |
| --- | --- | --- |
| **Caption** | Paris | Paris |
| **Name** | [Europe].[France].[Paris]（只读） | Paris |
| **SourceName** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同；只读） |
| **Value** | [Europe].[France].[Paris]（只读） | Paris |

在将一个索引值指定到 **[PivotItems](./../../PivotItems/PivotItems 对象.html)** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源） | 语法（非 OLAP 数据源） |
| --- | --- |
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris") |

在使用 **Item**属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源） | 名称（非 OLAP 数据源） |
| --- | --- |
| [Europe].[France].[Paris] | Paris |

## 示例

```javascript
/*本示例显示包含活动单元格的字段中第二个项目的标签文本。*/
function test() {
    console.log(ActiveCell.PivotField.PivotItems(2).Caption)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的字段“name”的第一项的标签文本。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems(1)
    console.log(pvtItem.Caption)
}
```
