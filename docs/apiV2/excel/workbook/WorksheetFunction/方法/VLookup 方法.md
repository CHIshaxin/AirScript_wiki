---
title: "VLookup 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/VLookup%20%E6%96%B9%E6%B3%95.html
---
# VLookup (方法)

在表格数组的首列查找值，并由此返回表格数组当前行中其他列的值。

## 说明

**VLookup** 中的 V 表示垂直方向。当比较值位于需要查找的数据左边的一列时，可以使用 **VLookup** 方法，而不用 **HLookup** 方法。

- 如果 Col_index_num 小于 1，则 **VLookup** 方法将生成错误。
- 如果 Col_index_num 大于 table_array 中的列数，则 **VLookup** 方法将生成错误。
- 如果 Range_lookup 为 TRUE 或被省略，则返回精确匹配值或近似匹配值。如果找不到精确匹配值，则返回下一个小于 lookup_value 的最大值。 table_array 第一列中的值必须以升序排序；否则 **VLookup** 方法可能无法给出正确的值。
- 如果 Range_lookup 为 FALSE，**VLookup** 方法将只查找精确匹配值。在此情况下，table_array 第一列中的值不需要排序。如果 table_array 第一列中有两个或更多个值与 lookup_value 匹配，则使用找到的第一个值。如果找不到精确匹配值，则生成错误。
- 在 table_array 第一列中搜索文本值时，请确保 table_array 第一列中的数据没有前置空格、尾部空格、直引号（' 或 "）与弯引号（‘ 或 “）不一致或非打印字符。在上述情况下，**VLookup** 方法可能给出不正确或非预期的值。有关如何清除或剪裁值的信息，请参阅 **Clean** 和 **Trim** 方法。
- 在搜索数字或日期值时，请确保 table_array 第一列中的数据未存储为文本值。在此情况下，**VLookup** 方法可能给出不正确或非预期的值。
- 如果 range_lookup 为 FALSE，而 lookup_value 为文本，则可以在 lookup_value 中使用通配符，包括问号 (?) 和星号 (\*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Lookup_value - 要在表格数组 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）第一列中搜索的值。Lookup_value 可以为值或引用。如果 lookup_value 小于 table_array 第一列中的最小值，则 VLOOKUP 将返回错误值 #N/A。 |
| Arg2 | any | 必填 | Table_array - 两列或更多列数据。可以使用对区域或区域名称的引用。table_array 第一列中的值是由 lookup_value 搜索的值。这些值可以是文本、数字或逻辑值。不区分大小写。 |
| Arg3 | any | 必填 | Col_index_num - table_array 中待返回的匹配值的列号。col_index_num 为 1 时，返回 table_array 第一列中的值；col_index_num 为 2 时，返回 table_array 第二列中的值，依此类推。 |
| Arg4 | any | 可选 | Range_lookup - 一个逻辑值，指定希望 **VLookup** 方法查找精确的匹配值还是近似匹配值。 |

## 返回值

Variant

## 示例

```javascript
/*本示例演示给定一组数据，用 VLookup 在数组的首列查找值，并由此返回表格数组当前行中其他列的值，并分别赋值到A3等单元格。*/
function test() {
    let arg2 = [[1, 2, 3],[4, 5, 6],[7, 8, 9]]
    Range("A3").Value2 = WorksheetFunction.VLookup(3, arg2, 2)
    Range("A4").Value2 = WorksheetFunction.VLookup(4, arg2, 1, false)
    Range("A5").Value2 = WorksheetFunction.VLookup(10, arg2, 1, true)
}
```

```javascript
/*本示例演示为A3等单元格赋值，用 VLookup 在表格数组的首列查找值，并由此返回表格数组当前行中第二列的值。*/
function test() {
    Range("A3").Value2 = 1.2
    Range("A4").Value2 = 2.3
    Range("A5").Value2 = 3.4
    Range("B3").Value2 = 5.1
    Range("B4").Value2 = 6.2
    Range("B5").Value2 = 7
    Range("D1").Value2 = 2.23
    console.log(WorksheetFunction.VLookup(Range("D1").Value2, Range("A3:B5"), 2))
}
```
