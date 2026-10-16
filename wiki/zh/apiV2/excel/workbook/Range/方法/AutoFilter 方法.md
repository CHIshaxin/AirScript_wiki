---
title: "AutoFilter (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/AutoFilter%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DcR16rBf
---

# AutoFilter (方法)

使用“自动筛选”筛选一个列表。

## 说明

如果忽略全部参数，此方法仅在指定区域切换自动筛选下拉箭头的显示。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Field | any | 可选 | 相对于作为筛选基准字段（从列表左侧开始，最左侧的字段为第一个字段）的字段的整型偏移量。 |
| Criteria1 | any | 可选 | 筛选条件（一个字符串；例如，“101”）。使用“=”可查找空字段，或者使用“<>”查找非空字段。如果省略该参数，则搜索条件为 All。如果将 *Operator* 设置为 **xlTop10Items**，则 *Criteria1* 指定数据项个数（例如，“10”）。 |
| Operator | Default | 可选 | 指定筛选类型的 XlAutoFilterOperator 常量之一。 |
| Criteria2 | any | 可选 | 第二个筛选条件（一个字符串）。与 *Criteria1* 和 *Operator* 一起组合成复合筛选条件。 |
| VisibleDropDown | any | 可选 | 如果为 **True**，则显示筛选字段的自动筛选下拉箭头。如果为 **False**，则隐藏筛选字段的自动筛选下拉箭头。默认值为 **True**。 |

## 返回值

Variant

## 示例

```javascript
/*本示例从工作表 Sheet1 上的 A1 单元格开始筛选出一个清单，该清单中只显示字段一为字符串“Otis”的数据项。字段一的下拉箭头将被隐藏。*/
function test() {
    Worksheets("Sheet1").Range("A1").AutoFilter(1, "Otis", null, null, false)
}
```

```javascript
/*本示例从工作表 Sheet1 上的 A1 单元格开始筛选出一个清单，该清单中显示字段一为字符串“Otis” 和 “4” 的数据项。*/
function test() {
    Worksheets("Sheet1").Range("A1").AutoFilter(1, "Otis", xlOr, "4", true)
}
```
