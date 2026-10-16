---
title: "Lookup 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Lookup%20%E6%96%B9%E6%B3%95.html
---
# Lookup (方法)

从单行或单列区域或者从一个数组返回值。LOOKUP 函数具有两种语法形式：向量形式和数组形式。

## 说明

| 如果需要 | 则参阅 | 用法 |
| --- | --- | --- |
| 在单行区域或单列区域（称为“向量”）中查找值，然后返回第二个单行区域或单列区域中相同位置的值。 | 向量形式 | 当要查询的值列表较大或者值可能会随时间而改变时，使用该向量形式。 |
| 在数组的第一行或第一列中查找指定的数值，然后返回数组的最后一行或最后一列中相同位置的数值。 | 数组形式 | 当值列表较小并且值在一段时间内保持不变时，使用该数组形式。 |

对于详细的测试或七种以上条件的测试，还可以使用 LOOKUP 函数作为 IF 函数的替代函数。请参阅数组形式的示例。

# 向量形式

向量是只包含一行或一列的区域。LOOKUP 的向量形式在单行区域或单列区域（称为“向量”）中查找值，然后返回第二个单行区域或单列区域中相同位置的值。当要指定的区域包含要匹配的值时，请使用 LOOKUP 函数的这种形式。LOOKUP 的另一种形式自动在第一行或第一列中查找。

**要点** lookup_vector 中的值必须以升序顺序放置：...,-2, -1, 0, 1, 2, ...；A-Z；FALSE, TRUE。否则，LOOKUP 可能无法提供正确的值。大写文本和小写文本是等同的。

- 如果 LOOKUP 找不到 lookup_value，则它与 lookup_vector 中小于或等于 lookup_value 的最大值匹配。
- 如果 lookup_value 小于 lookup_vector 中的最小值，则 LOOKUP 会提供 #N/A 错误值。

# 数组形式

LOOKUP 的数组形式在数组 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）的第一行或第一列中查找指定的值，并返回数组最后一行或最后一列内相同位置的值。当要匹配的值位于数组的第一行或第一列中时，使用 LOOKUP 的这种形式。当要指定列或行的位置时，使用 LOOKUP 的另一种形式。

**提示** 一般而言，最好使用 HLOOKUP 或 VLOOKUP 函数而不是 LOOKUP 的数组形式。LOOKUP 的这种形式是为了与其他电子表格程序兼容而提供的。

- 如果 LOOKUP 找不到 lookup_value，它会使用数组中小于或等于 lookup_value 的最大值。
- 如果 lookup_value 小于第一行或第一列中的最小值（取决于数组维度），LOOKUP 会返回 #N/A 错误值。

LOOKUP 的数组形式与 HLOOKUP 和 VLOOKUP 函数非常相似。区别在于：HLOOKUP 在第一行中搜索 lookup_value，VLOOKUP 在第一列中搜索，而 LOOKUP 根据数组维度进行搜索。

- 如果数组包含宽度比高度大的区域（列数多于行数），LOOKUP 会在第一行中搜索 lookup_value。
- 如果数组是正方的或者高度大于宽度（行数多于列数），则 LOOKUP 在第一列中进行搜索。
- 使用 HLOOKUP 和 VLOOKUP，可以通过索引以向下或遍历的方式搜索，但是 LOOKUP 始终选择行或列中的最后一个值。

**要点** 数组中的值必须以升序顺序放置：...,-2, -1, 0, 1, 2, ...；A-Z；FALSE, TRUE。否则，LOOKUP 无法提供正确的值。大写文本和小写文本是等同的。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Lookup_value - LOOKUP 在第一个向量中搜索的值。Lookup_value 可以是数字、文本、逻辑值、名称或对值的引用。 |
| Arg2 | any | 必填 | Lookup_vector 或数组 - 在向量形式中，为只包含一行或一列的区域。lookup_vector 中的值可以是文本、数字或逻辑值。在数组形式中，为包含要与 lookup_value 进行比较的文本、数字或逻辑值的单元格区域。 |
| Arg3 | any | 可选 | Result_vector - 只用于向量形式。只包含一行或一列的区域。它必须与 lookup_vector 大小相同。 |

## 返回值

Variant

## 示例

```javascript
/*本示例使用 Lookup 方法从一个数组获取值，并分别赋值给B2和B3单元格。*/
function test() {
    let arr1 = [1, 2, 3]
    let arr2 = ["苹果", "香蕉", "西瓜"]
    Range("B2").Value2 = WorksheetFunction.Lookup(2, arr1, arr2)
    Range("B3").Value2 = WorksheetFunction.Lookup("西瓜", arr2, arr1)
}
```

```javascript
/*本示例为D1等单元格分别赋值，使用 Lookup 方法从单列区域D1:D5获取值。*/
function test() {
    Range("D1").Value2 = 5
    Range("D2").Value2 = 4
    Range("D3").Value2 = -30
    Range("D4").Value2 = 10
    Range("D5").Value2 = 1.3
    console.log(WorksheetFunction.Lookup(0, Range("D1:D5")))
}
```
