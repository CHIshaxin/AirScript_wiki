---
title: "Match 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Match%20%E6%96%B9%E6%B3%95.html
---
# Match (方法)

返回在指定方式下与指定数值匹配的数组 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）中元素的相应位置。如果需要找出匹配元素的位置而不是匹配元素本身，则应该使用 MATCH 函数而不是 LOOKUP 函数。

## 说明

- Lookup_value 为需要在 Look_array 中查找的数值。例如，如果要在电话簿中查找某人的电话号码，则应该将姓名作为查找值，但实际上需要的是电话号码。
- Lookup_value 可以为值（数字、文本或逻辑值）或对数字、文本或逻辑值的单元格引用。
- 如果 match_type 为 1，函数 MATCH 查找小于或等于 lookup_value 的最大数值。Lookup_array 必须按升序排列：...、-2、-1、0、1、2、...、A-Z、FALSE、TRUE。
- 如果 match_type 为 0，函数 MATCH 查找等于 lookup_value 的第一个数值。Lookup_array 可以按任何顺序排列。
- 如果 match_type 为 -1，函数 MATCH 查找大于或等于 lookup_value 的最小数值。Lookup_array 必须按降序排列：TRUE、FALSE、Z-A、 ...2、 1、 0、 -1、 -2、 ...，等等。
- 如果省略 match_type，则假设为 1。
- 函数 MATCH 返回 lookup_array 中目标值的位置，而不是数值本身。例如，MATCH("b",{"a","b","c"},0) 返回 2，即“b”在数组 {"a","b","c"} 中的相应位置。
- 匹配文本值时，函数 MATCH 不区分大小写字母。
- 如果函数 MATCH 未能找到匹配值，则返回错误值 #N/A。
- 如果 match_type 为 0 且 lookup_value 为文本，则可以在 lookup_value 中使用通配符、问号 (?) 和星号 (\*)。问号匹配任意单个字符；星号匹配任意字符序列。如果您要查找实际的问号或星号本身，请在该字符前键入波形符 (~)。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Lookup_value - 需要在表中查找的值。 |
| Arg2 | any | 必填 | Lookup_array - 可能包含所要查找的值的连续单元格区域。Lookup_array 必须为数组或数组引用。 |
| Arg3 | any | 可选 | Match_type - 数字 -1、0 或 1。Match_type 指明 ET 如何将 lookup_value 与 lookup_array 中的值进行匹配。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Match 方法查找数组arr1中小于等于7的最大数值和等于3的第一个数值的对应位置，并分别赋值给A1和A2单元格。*/
function test() {
    let arr1 = [1, 2, 3, 4, 5]
    Range("A1").Value2 = WorksheetFunction.Match(7, arr1, 1)
    Range("A2").Value2 = WorksheetFunction.Match(3, arr1, 0)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Match 方法查找大于等于322的最小数值的相应位置。*/
function test() {
    Range("B1").Value2 = 500
    Range("B2").Value2 = 412
    Range("B3").Value2 = -30
    Range("B4").Value2 = 230
    Range("B5").Value2 = 1.3
    console.log(WorksheetFunction.Match(322, Range("B1:B5"), -1))
}
```
