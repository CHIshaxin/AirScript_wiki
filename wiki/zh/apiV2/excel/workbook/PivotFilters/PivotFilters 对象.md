---
title: "PivotFilters (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilters/PivotFilters%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Chc6AtDN
---

# PivotFilters (对象)

**[PivotFilters](./../PivotFilters/PivotFilters 对象.html)** 对象是 **[PivotFilter](./../PivotFilter/PivotFilter 对象.html)** 对象的集合。

## 说明

**[PivotFilters](./../PivotFilters/PivotFilters 对象.html)** 集合包含用于添加新筛选器、计算集合中现有筛选器个数以及引用特定 **[PivotFilter](./../PivotFilter/PivotFilter 对象.html)** 对象的属性和方法。

## 示例

```javascript
/*本示例为活动工作表的数据透视表的的字段“date”添加筛选器。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("date").PivotFilters.Add2(xlDateThisWeek)
}
```

```javascript
/*本示例显示工作表 Sheet1 中数据透视表字段“name”的第一个筛选器中作为筛选依据的值字段。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.DataField.Value)
}
```
