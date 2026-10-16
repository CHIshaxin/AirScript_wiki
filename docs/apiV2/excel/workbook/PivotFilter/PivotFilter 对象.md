---
title: "PivotFilter 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/PivotFilter%20%E5%AF%B9%E8%B1%A1.html
---
# PivotFilter (对象)

PivotFilter 应用于 **[PivotField](./../PivotField/PivotField 对象.html)** 对象。

## 说明

开发人员可以选择命名筛选器以供引用，因为索引不可靠。 **DataField** 属性指定要基于值筛选器的 PivotField。

## 示例

```javascript
/*本示例显示第一张工作表中数据透视表的字段“id”的第一个筛选器的参数。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id")
    console.log(pvtField.PivotFilters.Item(1).Value1)
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“score”的第二个筛选器是否是活动的。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("score").PivotFilters.Item(2).Active)
}
```
