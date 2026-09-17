---
title: "Characters 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Characters%20%E6%96%B9%E6%B3%95.html
---
# Characters (方法)

返回 **[Characters](./../../Characters/Characters 对象.html)** 对象，它代表对象文本内某个区域的字符。使用 **[Characters](./../../Characters/Characters 对象.html)** 对象可为文本字符串内的字符设置格式。

## 说明

**[Characters](./../../Characters/Characters 对象.html)** 对象不是集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Start | any | 可选 | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。 |
| Length | any | 可选 | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（*Start* 字符之后的所有字符）。 |

## 示例

```javascript
/*此示例将 Sheet1 中 A1 单元格第三个字符的格式设为加粗。*/
function test() {
    let a1 = Worksheets.Item("Sheet1").Range("A1")
    a1.Value(undefined, "abcdefg")
    a1.Characters(3, 1).Font.Bold = true
}
```

```javascript
/*本示例将 A1 单元格的值设置为 This is A1,并显示 A1 单元格前四个字符的文本*/
function test() {
    Range("A1").Value2 = "This is A1"
    console.log(Range("A1").Characters(1, 4).Text)
}
```
