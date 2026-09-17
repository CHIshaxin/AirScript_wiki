---
title: "ClearTable 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/ClearTable%20%E6%96%B9%E6%B3%95.html
---
# ClearTable (方法)

**ClearTable**方法用于清除数据透视表。清除数据透视表包括删除所有字段以及删除应用于数据透视表的所有筛选和排序。此方法将数据透视表重置为该表刚创建而尚未添加任何字段时的状态。

## 说明

**ClearTable**函数不使用任何参数，可用于关系数据透视表和 OLAP 数据透视表。

## 示例

```javascript
/*以下示例清除活动工作表中的数据透视表。*/
function test() {
    ActiveSheet.PivotTables(1).ClearTable()
}
```

```javascript
/*本示例删除活动工作表中所有的数据透视表*/
function test() {
    for (i = 1; i <= ActiveSheet.PivotTables().Count; i++) {
        ActiveSheet.PivotTables(i).ClearTable()
    }
}
```
