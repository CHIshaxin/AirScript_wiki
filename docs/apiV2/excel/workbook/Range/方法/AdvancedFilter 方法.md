---
title: "AdvancedFilter 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/AdvancedFilter%20%E6%96%B9%E6%B3%95.html
---
# AdvancedFilter (方法)

基于条件区域从列表中筛选或复制数据。如果初始选定区域为单个单元格，则使用单元格的当前区域。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Action | XlFilterAction | 必填 | **[XlFilterAction](./../../枚举/XlFilterAction 枚举.html)** 的常量之一，用于指定是否就地复制或筛选列表。 |
| CriteriaRange | any | 可选 | 条件区域。如果省略该参数，则没有条件限制。 |
| CopyToRange | any | 可选 | 如果 *Action* 为 **xlFilterCopy**，则为复制行的目标区域。否则，忽略该参数。 |
| Unique | any | 可选 | 如果为 **True**，则只筛选唯一记录。如果为 **False**，则筛选符合条件的所有记录。默认值为 **False**。 |

## 返回值

Variant

## 示例

```javascript
/*本示例基于条件区域“Criteria”筛选数据库“Database”*/
function test() {
    Range("Database").AdvancedFilter(xlFilterInPlace, Range("Criteria"))
}
```

```javascript
/*本示例将 A1:C3 单元格区域以 E9:G10 单元格区域为条件复制到 I1 单元格中。*/
function test() {
    Range("A1:C3").AdvancedFilter(xlFilterCopy, Range("E9:G10"), Range("I1"), false)
}
```
