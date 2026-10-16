---
title: "AutoSort (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/AutoSort%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B6d-PKCk
---

# AutoSort (方法)

为数据透视表建立自动字段排序规则。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Order | int4 | 必填 | 指定排序顺序的 XlSortOrder 的常量之一。 |
| Field | string | 必填 | 排序关键字段的名称。必须指定唯一名称（由 SourceName 属性返回的名称），而不是显示的名称。 |
| PivotLine | any | 可选 | 列上的一行或数据透视表中的一行。 |
| CustomSubtotal | any | 可选 | 自定义分类汇总字段。 |

## 示例

```javascript
/*本示例对活动工作表中第一个数据透视表的“grades”字段按降序排序。*/
function test() {
    Application.ActiveSheet.PivotTables(1).PivotFields("grades").AutoSort(xlDescending, "grades")
}
```

```javascript
/*本示例对工作表 Sheet1 中第一张数据透视表的字段“id”按升序排序。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("id")
    pvtField.AutoSort(xlAscending, "id")
}
```
