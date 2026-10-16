---
title: "SourceData (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/SourceData%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: qOUP41y4
---

# SourceData (属性)

返回数据透视表的数据源，如下表所示。**Variant** 型，可读写。

## 说明

| 数据源 | 返回值 |
| --- | --- |
| ET 列表或数据库 | 以文本形式返回的单元格引用。 |
| 外部数据源 | 数组。每行包含一个 SQL 连接字符串，该行的剩余元素作为查询字符串，以 255 个字符为单位分段。 |
| 多个合并区域 | 二维数组。每行由一个引用及与其相关联的页字段项组成。 |
| 其他数据透视表 | 以上三种情形之一。 |

该属性不可用于 OLE DB 数据源。

## 示例

```javascript
/*本示例显示第一张工作表上的第一个数据透视表的数据源。*/
function test() {
    console.log(Application.Worksheets.Item(1).PivotTables(1).SourceData)
}
```

```javascript
/*本示例设置活动工作表上数据透视表的数据源，并刷新该数据透视表。*/
function test() {
    let pvtTable = Application.ActiveSheet.Range("I1").PivotTable
    pvtTable.SourceData = "=Sheet1!R1C1:R9C3"
    pvtTable.RefreshTable()
}
```
