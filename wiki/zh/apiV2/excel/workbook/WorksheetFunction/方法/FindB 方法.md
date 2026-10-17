---
title: "FindB (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/FindB%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: WaZLljqo
---

# FindB (方法)

FIND 与 FINDB 查找一个文本字符串内的另一个文本字符串，并返回一个数字，表示从文本字符串（从中查找另一个文本字符串）的第一个字符起开始查找的第一个文本字符串的起始位置。

## 说明

FIND 用于使用单字节字符集 (SBCS) 的语言，而 FINDB 用于使用双字节字符集 (DBCS) 的语言。计算机上的默认语言设置会以下列方式影响返回值：

- 无论默认语言设置如何，FIND 始终将每个字符（不管是单字节还是双字节字符）作为 1 个字符计算。
- 启用对支持 DBCS 的语言的编辑并将它设置为默认语言时，FINDB 会将每个双字节字符作为 2 个字符计算。否则 FINDB 会将每个字符作为 1 个字符计算。

支持 DBCS 的语言包括日语、中文（简体）、中文（繁体）和朝鲜语。

- FIND 和 FINDB 区分大小写并且不允许使用通配符。如果不想执行区分大小写的搜索或要使用通配符，可以使用 SEARCH 和 SEARCHB。
- 如果 find_text 为 ""（空文本），在 FIND 将匹配搜索字符串中的第一个字符（即，编号为 start_num 或 1 的字符）。
- Find_text 不能包含任何通配符。
- 如果 within_text 中没有 find_text，则 FIND 和 FINDB 将返回错误值 #VALUE!。
- 如果 start_num 不大于 0，则 FIND 和 FINDB 将返回错误值 #VALUE!。
- 如果 start_num 大于 within_text 的长度，则 FIND 和 FINDB 将返回错误值 #VALUE!。
- 使用 start_num 可跳过指定编号的字符。以 FIND 为例，假设要处理文本字符串“AYF0093.YoungMensApparel”。要在文本字符串的说明部分中查找第一个“Y”的编号，请将 start_num 设置为等于 8，这样就不会搜索文本的序列号部分。FIND 从第 8 个字符开始，在下一个字符处查找 find_text，并返回数字 9。FIND 总是返回从 within_text 的起始位置计算的字符的编号，如果 start_num 大于 1，则会计算跳过的字符。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | Find_text - 要查找的文本。 |
| Arg2 | string | 必填 | Within_text - 包含要查找的文本的文本。 |
| Arg3 | any | 可选 | Start_num - 指定要从其开始搜索的字符。within_text 中的第一个字符是编号为 1 的字符。如果省略 start_num，则假定其值为 1。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值A1等单元格，使用 FindB 方法查找参数1字符串开始的位置。*/
function test() {
    Range("A1").Value2 = "wps"
    Range("A2").Value2 = "hello, wps!"
    Range("A3").Value2 = 2
    console.log(WorksheetFunction.FindB(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```

```javascript
/*本示例使用 FindB 方法查找参数1字符串开始的位置，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.FindB("are", "how are you")
    Range("B2").Value2 = WorksheetFunction.FindB("you", "Will you designate the flowers you wish", 7)
}
```
