---
title: "Characters 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Characters/Characters%20%E5%AF%B9%E8%B1%A1.html
---
# Characters (对象)

代表包含文本的对象中的字符。

## 说明

使用 **[Characters](./../Characters/Characters 对象.html)** 对象可修改包含在全文本字符串中的任意字符序列。

使用 **Characters**(*start*, *length* )（其中 *start* 为起始字符号，而 *length* 为要返回的字符个数）返回 **[Characters](./../Characters/Characters 对象.html)** 对象。

```javascript
/*下例向单元格 B1 中添加文本，并将第二个单词设置为加粗。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("B1").Value = "New Title"
    Application.Worksheets.Item("Sheet1").Range("B1").Characters(5, 5).Font.Bold = true
}
```

仅当需要更改对象中文本的一部分而不影响其余部分时，才有必要使用 **Characters** 方法（如果对象不支持格式文本，则不能使用 **Characters** 方法对文本中的一部分单独设置格式）。要同时更改所有文本，通常可以对该对象直接应用某一适当的方法或属性。

```javascript
/*下例将单元格 A5 的内容设置为倾斜。*/
Application.Worksheets.Item("Sheet1").Range("A5").Font.Italic = true
```
