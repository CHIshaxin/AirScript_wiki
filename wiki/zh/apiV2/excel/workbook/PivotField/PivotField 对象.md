---
title: "PivotField (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/PivotField%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BhE0WnF3
---

# PivotField (对象)

代表数据透视表中的一个字段。

## 说明

**[PivotField](./../PivotField/PivotField 对象.html)** 对象是 **[PivotFields](./../PivotFields/PivotFields 对象.html)** 集合的成员。 **[PivotFields](./../PivotFields/PivotFields 对象.html)** 集合包含数据透视表中的所有字段，包括隐藏字段。

下列属性返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

- **ColumnFields** 属性
- **DataFields** 属性
- **HiddenFields** 属性
- **PageFields** 属性
- **RowFields** 属性
- **VisibleFields** 属性

使用 **PivotFields**( *index* )（其中 *index* 是字段名称或索引号）可返回一个 **[PivotField](./../PivotField/PivotField 对象.html)** 对象。

## 示例

```javascript
/*本示例使 Sheet1 上第一张数据透视表中的字段“id”成为行字段。*/
function test() {
    Application.Worksheets.Item("Sheet1").PivotTables(1).PivotFields("id").Orientation = xlRowField
}
```

```javascript
/*本示例删除工作表 Sheet1 中数据透视表应用于字段“score”的所有筛选。*/
function test() {
    let pvtfield = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("score")
    pvtfield.ClearAllFilters()
}
```
