---
title: "DataRange (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/DataRange%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DW-oaamj
---

# DataRange (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，如下表所示。只读。

## 说明

| 对象 | 数据区域 |
| --- | --- |
| “数据”字段 | 字段中包含的数据 |
| 行、列或者页字段 | 字段中的数据项 |
| 数据项 | 数据项规范所限定的数据 |

## 示例

```javascript
/*本示例选中工作表 Sheet1 中数据透视表字段“name”的数据透视表项。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    pvtTable.PivotFields("name").DataRange.Select()
}
```

```javascript
/*本示例显示活动工作表中数据透视表字段“grades”的数据透视表项。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("grades")
    console.log(pvtField.DataRange.Value2.toString())
}
```
