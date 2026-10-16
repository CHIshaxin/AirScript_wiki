---
title: "Subtotals (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/Subtotals%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D09WJJav
---

# Subtotals (方法)

返回或设置与指定字段同时显示的分类汇总。仅对非数据字段有效。**Variant** 类型，可读写。

## 说明

如果参数 Index 为 **True**，则该字段显示分类汇总。如果 Index 传 1（自动）为 **True**，则其他所有值将设置为 **False**。

| Index | 含义 |
| --- | --- |
| 1 | 自动 |
| 2 | 总计 |
| 3 | 计数 |
| 4 | 平均 |
| 5 | 最大值 |
| 6 | 最小值 |
| 7 | 乘积 |
| 8 | Count Nums |
| 9 | StdDev |
| 10 | StdDevp |
| 11 | Var |
| 12 | Varp |

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，*Index* 只能返回或设置为 1（自动）。返回的数组对第一个数组元素总是包含 **True** 或 **False**，而对所有其他元素则包含 **False**。数组中所有元素的值都为 **False**，则表明没有分类汇总。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 表示下表所显示的分类汇总索引。如果省略此参数，则 Subtotals 方法将返回包含每个分类汇总的相应 Boolean 值的数组。 |
| RHS | any | 可选 | 当设置分类汇总时，此参数用于指定分类汇总的相应 Boolean 值。若要返回分类汇总结果，无须传此参数。 |

## 示例

```javascript
/*本示例将包含活动单元格的字段设置为显示分类汇总求和。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    ActiveCell.PivotField.Subtotals(2, true)
}
```

```javascript
/*本示例将活动工作表第一张数据透视表的字段“name”显示的分类汇总设置为“Max”，并显示是否设置成功。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("name")
    pvtField.Subtotals(5, true)
    console.log(pvtField.Subtotals(5))
}
```
