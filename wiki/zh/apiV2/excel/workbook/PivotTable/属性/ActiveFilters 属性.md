---
title: "ActiveFilters (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ActiveFilters%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B2SggJrr
---

# ActiveFilters (属性)

指示指定数据透视表中的当前活动的筛选。只读。

## 示例

```javascript
/*本示例显示 A2 单元格所在数据透视表活动筛选的数量*/
function test() {
    console.log(Range("A2").PivotTable.ActiveFilters.Count)
}
```

```javascript
/*本示例显示 A2 单元格所在数据透视表的活动筛选中的最后一项的Value1*/
function test() {
    let filters = Range("A2").PivotTable.ActiveFilters
    console.log(filters.Item(filters.Count).Value1)
}
```
