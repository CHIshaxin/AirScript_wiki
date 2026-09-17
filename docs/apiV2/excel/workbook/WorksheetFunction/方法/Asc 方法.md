---
title: "Asc 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Asc%20%E6%96%B9%E6%B3%95.html
---
# Asc (方法)

对于双字节字符集 (DBCS) 语言，将全角（双字节）字符更改为半角（单字节）字符。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | 文本或对包含要更改的文本的单元格的引用。如果文本中不包含任何全角字母，则文本不会更改。 |

## 返回值

String

## 示例

```javascript
/*本示例为C3等单元格分别赋值，并使用Asc方法将全角字符更改为半角字符。*/
function test() {
    Range("C3").Value2 = "Hello，World！"
    Range("C4").Value2 = "Hi，How are you?"
    Range("C5").Value2 = "Fine，thank you！"
    let asc1 = Application.WorksheetFunction.Asc(Range("C3").Value2)
    let asc2 = Application.WorksheetFunction.Asc(Range("C4").Value2)
    let asc3 = Application.WorksheetFunction.Asc(Range("C5").Value2)
    console.log(asc1)
    console.log(asc2)
    console.log(asc3)
}
```

```javascript
/*本示例使用Asc方法将全角字符更改为半角字符，并将结果赋值于J3等单元格。*/
function test() {
    Range("J3").Value2 = Application.WorksheetFunction.Asc("apple，banana，pear。")
    Range("J4").Value2 = Application.WorksheetFunction.Asc("red；blue；purple。")
    Range("J5").Value2 = Application.WorksheetFunction.Asc("mark：Hello！")
}
```
