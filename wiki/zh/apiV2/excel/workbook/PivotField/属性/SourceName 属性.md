---
title: "SourceName (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/SourceName%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B0i6jmuF
---

# SourceName (属性)

返回一个 **String** 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。

## 说明

如果用户在创建数据透视表之后重命名数据项，此属性的值可能会与当前数据项名称不同。

下表显示 **SourceName** 属性及其相关属性的示例值，假设存在唯一名称为“[Europe].[France].[Paris]”的 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在包含名为“Paris”的项的非 OLAP 数据源。

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
/*本示例显示第一张工作表中第一张数据透视表的字段“name”的原始源数据中展示的名称。*/
function test() {
    console.log(Worksheets.Item(1).PivotTables(1).PivotFields("name").SourceName)
}
```

```javascript
/*本示例显示第一张工作表中数据透视表的字段“求和项:score”的原始源数据中展示的名称。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("求和项:score")
    console.log(pvtField.SourceName)
}
```
