---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DlOcErJh
---

# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 说明

下表显示了 **Name** 属性和相关属性的示例值，同时假设有一个具有唯一名称为“[Europe].[France].[Paris]”的 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，有一个包含项名称为“Paris”的非 OLAP 数据源。

| 属性 | 值（OLAP 数据源） | 值（非 OLAP 数据源） |
| --- | --- | --- |
| **Caption** | Paris | Paris |
| **Name** | [Europe].[France].[Paris]（只读） | Paris |
| **SourceName** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同，只读） |
| **Value** | [Europe].[France].[Paris]（只读） | Paris |

在将一个索引值指定到 **[PivotItems](./../../PivotItems/PivotItems 对象.html)** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源） | 语法（非 OLAP 数据源） |
| --- | --- |
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源） | 名称（非 OLAP 数据源） |
| --- | --- |
| [Europe].[France].[Paris] | Paris |

## 示例

```javascript
/*本示例显示活动工作表中数据透视表的字段“name”的第一个数据项的名称。*/
function test() {
    let pvtItem = ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotItems(1)
    console.log(pvtItem.Name)
}
```

```javascript
/*本示例将工作表 Sheet1 上第一张数据透视表的字段“id”中第一个数据项的名称设置为“9”。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("id").PivotItems(1)
    pvtItem.Name = "9"
}
```
