---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilters/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

根据 **[PivotFilters](./../../PivotFilters/PivotFilters 对象.html)** 集合对象特定元素在集合中的位置返回该特定元素。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 指定集合的元素的位置。 |

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表字段“score”的第一个筛选器应用的筛选器类型是否为“xlCaptionIsGreaterThan”。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("score").PivotFilters.Item(1)
    console.log(pvtFilter.FilterType == xlCaptionIsGreaterThan)
}
```

```javascript
/*本示例显示第一张工作表中数据透视表的字段“id”的第一个筛选器的参数。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id")
    console.log(pvtField.PivotFilters.Item(1).Value1)
}
```
