---
title: "Replace 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Replace%20%E6%96%B9%E6%B3%95.html
---
# Replace (方法)

根据指定的字符数，将文本字符串的部分用其他文本字符串替换。返回值是替换之后的新字符串。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | 要替换其中某些字符的文本。 |
| Arg2 | double | 必填 | 在 **Arg1** 中用 **Arg4** 替换的字符的位置。 |
| Arg3 | double | 必填 | 在 **Arg1** 中使用 **Replace** 方法替换为 **Arg4** 的字符个数。 |
| Arg4 | string | 必填 | 用以替换 **Arg1** 中字符的文本。 |

## 返回值

String

## 示例

```javascript
/*此示例用“ac-ef”替换“abcdef”，并在该过程中通知用户。*/
function test() {
    let strCurrent = "abcdef"
    // Notify user and display current string.
    console.log("The current string is:" + strCurrent)
    //Replace "cd" with "-".
    let strReplaced = Application.WorksheetFunction.Replace (strCurrent, 3, 2, "-")
    // Notify user and display replaced string.
    console.log ("The replaced string is:" + strReplaced)
}
```

```javascript
/*本示例为A1等单元格赋值，使用 Replace 方法根据指定的字符数将文本字符串的一部分替换为其他文本字符串。*/
function test() {
    Range("A1").Value2 = "欢迎你加入WPS大家庭！"
    Range("A2").Value2 = 6
    Range("A3").Value2 = 3
    Range("A4").Value2 = "金山办公"
    console.log(WorksheetFunction.Replace(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2, Range("A4").Value2))
}
```
