---
title: "Add2 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilters/%E6%96%B9%E6%B3%95/Add2%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: U7TmUWZ3
---

# Add2 (方法)

将新筛选添加到 **[PivotFilters](./../../PivotFilters/PivotFilters 对象.html)** 集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | XlPivotFilterType | 必填 | 要求 XlPivotFilterType 类型的筛选。 |
| DataField | any | 可选 | 筛选附加到的字段。 |
| Value1 | any | 可选 | 筛选值 1。 |
| Value2 | any | 可选 | 筛选值 1。 |
| Order | any | 可选 | 筛选数据的顺序。 |
| Name | any | 可选 | 筛选的名称。 |
| Description | any | 可选 | 筛选的简要说明 |
| MemberPropertyField | any | 可选 | 指定标签筛选所基于的成员属性字段。 |

## 返回值

PivotFilter

## 示例

```javascript
/*本示例为活动工作表的数据透视表的字段“date”添加筛选器。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("date").PivotFilters.Add2(xlDateThisWeek)
}
```

```javascript
/*本示例为活动工作表的数据透视表的字段“name”添加值筛选器。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotFilters.Add2(xlValueIsGreaterThanOrEqualTo, ActiveSheet.PivotTables(1).DataFields(1), 95)
}
```

```javascript
/*本示例为活动工作表中数据透视表的字段“id”添加值筛选器。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("id").PivotFilters.Add2(xlTopCount, ActiveSheet.PivotTables(1).DataFields("求和项:score"), 3)
}
```

```javascript
/*本示例为活动工作表中数据透视表的字段“id”添加标签筛选器。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("id").PivotFilters.Add2(xlCaptionIsNotBetween, null, "1", "5")
}
```

```javascript
/*本示例为活动工作表的数据透视表的字段“name”添加值筛选器。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotFilters.Add2(xlValueIsGreaterThanOrEqualTo, ActiveSheet.PivotTables(1).DataFields("求和项:期中成绩"), 75.5)
}
```
