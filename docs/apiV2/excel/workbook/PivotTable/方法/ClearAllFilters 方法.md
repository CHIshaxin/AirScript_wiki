---
title: "ClearAllFilters 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/ClearAllFilters%20%E6%96%B9%E6%B3%95.html
---
# ClearAllFilters (方法)

**ClearAllFilters** 方法将删除当前应用于数据透视表的所有筛选。包括删除 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象的 **[PivotFilters](./../../PivotFilters/PivotFilters 对象.html)** 集合中的所有筛选，删除应用的所有手动筛选以及将报表筛选区域中的所有筛选字段设置为默认项。

## 说明

调用 **ClearAllFilters** 方法后， **[PivotFilters](./../../PivotFilters/PivotFilters 对象.html)** 集合将为空。

## 示例

```javascript
/*本示例删除数据透视表1中的的所有筛选*/
function test() {
    ActiveSheet.PivotTables("数据透视表1").ClearAllFilters()
}
```

```javascript
/*本示例删除活动工作表中所有数据透视表中的所有筛选*/
function test() {
    for (i = 1; i <= ActiveSheet.PivotTables().Count; i++) {
        ActiveSheet.PivotTables(i).ClearAllFilters()
    }
}
```
