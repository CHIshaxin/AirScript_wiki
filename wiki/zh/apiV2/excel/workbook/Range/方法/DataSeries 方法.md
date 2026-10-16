---
title: "DataSeries (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/DataSeries%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DV4y41Ok
---

# DataSeries (方法)

在指定区域内创建数据系列。**Variant** 类型。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Rowcol | any | 可选 | 可以是 **xlRows** 或 **xlColumns** 常量，分别表示按行或列输入数据系列。如果省略本参数，则使用区域的大小和形状。 |
| Date | Default | 可选 | 如果 *Type* 参数为 **xlChronological**，则 *Date* 参数指示单步执行日期单位。 |
| Step | any | 可选 | 系列的步长值。默认值为 1。 |
| Stop | any | 可选 | 系列的终止值。如果省略本参数，ET 将填满整个区域。 |
| Trend | any | 可选 | 如果为 **True**，则创建一个线性趋势或增长趋势。如果为 **False**，则创建一个标准数据序列。默认值为 **False**。 |

## 返回值

Variant

## 示例

```javascript
/*本示例在工作表 Sheet1 上的单元格区域 A1:A12 中创建包含 12 个日期的数据系列。该系列包含 1996 年中每个月份的最后一天。*/
function test() {
    let dateRange = Worksheets.Item("Sheet1").Range("A1:A12")
    Worksheets.Item("Sheet1").Range("A1").Formula = "31-JAN-1996"
    dateRange.DataSeries(undefined, xlChronological, xlMonth)
}
```

```javascript
/*本示例将 A1 单元格的值设置为 100.6，对 A1:H5 单元格区域进行步长为 20，终止值为 180 的按行填充*/
function test() {
    Range("A1").Value2 = 100.6
    Range("A1:H5").DataSeries(xlRows, xlDataSeriesLinear, null, 20, 180)
}
```
