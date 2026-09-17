---
title: "Characters 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E6%96%B9%E6%B3%95/Characters%20%E6%96%B9%E6%B3%95.html
---
# Characters (方法)

返回一个 **[Characters](./../../Characters/Characters 对象.html)** 对象，该对象表示某个形状的文本框架中的字符区域。可以使用 **[Characters](./../../Characters/Characters 对象.html)** 对象向文本框架中添加字符和设置字符的格式。

## 说明

**[Characters](./../../Characters/Characters 对象.html)** 对象不是集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Start | any | 可选 | 要返回的第一个字符。如果此参数设置为 1 或被省略，则 Characters 方法将返回以第一个字符为起始字符的字符区域。 |
| Length | any | 可选 | 要返回的字符个数。如果省略此参数，则 Characters 方法将返回该字符串的剩余部分（由 Start 参数设置的字符以后的所有字符）。 |

## 示例

```javascript
/*本示例将活动工作表中第一个形状的文本框架中第三个字符的格式设置为加粗。*/
function test() {
    let textFrame = Application.ActiveSheet.Shapes.Item(1).TextFrame
    textFrame.Characters().Text = "abcdefg"
    textFrame.Characters(3, 1).Font.Bold = true
}
```

```javascript
/*本示例显示第一张工作表中第二个形状的文本框架中字符的数量。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.Item(2).TextFrame
    textFrame.Characters().Text = "hello"
    console.log(textFrame.Characters().Count)
}
```
