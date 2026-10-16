---
title: "Active 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/Active%20%E5%B1%9E%E6%80%A7.html
---
# Active (属性)

返回指定透视筛选是否是活动的。只读 **Boolean**。

## 说明

当筛选的透视字段在数据透视表中，且数据透视表更新时会计算该筛选时，此属性返回 **True**。当筛选的透视字段不在数据透视表中，且对数据透视表计算没有影响时，此属性返回 **False**。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“score”的第二个筛选器是否为活动的。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("score").PivotFilters.Item(2).Active)
}
```

```javascript
/*本示例显示工作表 Sheet1 中数据透视表字段“name”的第一个筛选器是否为活动的。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.Active)
}
```
