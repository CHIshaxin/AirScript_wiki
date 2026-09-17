---
title: "Fixed 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Fixed%20%E6%96%B9%E6%B3%95.html
---
# Fixed (方法)

将数字舍入到指定的小数位数，使用句点和逗号，以十进制数格式对该数进行格式设置，并以文本形式返回结果。

## 说明

- 在 ET 中，Numbers 的最大有效位数不能超过 15 位，但 decimals 可达到 127 位。
- 如果 decimals 为负数，则 number 向小数点的左边舍入。
- 如果省略 decimals，则假定其值为 2。
- 如果 no_commas 为 FALSE 或被省略，则返回的文本会像通常一样包含逗号。
- 该函数与使用\*\*“格式”**菜单上的**“单元格”**命令格式化包含数字的单元格的主要区别在于：函数 FIXED 将其结果转换成文本，而**“单元格”\*\*命令格式化的结果仍是数字。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要进行舍入并转换为文本的数字。 |
| Arg2 | any | 可选 | Decimals - 小数点右边的位数。 |
| Arg3 | any | 可选 | No_commas - 一个逻辑值，如果为 TRUE，则会阻止 FIXED 在返回的文本中包含逗号。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Fixed 方法将数字舍入到指定的小数位数,并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Fixed(-3234.5678, 2, false)
    Range("B2").Value2 = WorksheetFunction.Fixed(32584.5678, 3, true)
    Range("B3").Value2 = WorksheetFunction.Fixed(5234.96783)
}
```

```javascript
/*本示例为E1等单元格分别赋值，使用 Fixed 方法将数字舍入到指定的小数位数。*/
function test() {
    Range("E1").Value2 = 0.12345
    Range("E2").Value2 = 3
    console.log(WorksheetFunction.Fixed(Range("E1").Value2, Range("E2").Value2, true))
}
```
