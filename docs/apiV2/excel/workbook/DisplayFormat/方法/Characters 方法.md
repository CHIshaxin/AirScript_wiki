---
title: "Characters 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E6%96%B9%E6%B3%95/Characters%20%E6%96%B9%E6%B3%95.html
---
# Characters (方法)

返回一个 **[Characters](./../../Characters/Characters 对象.html)** 对象，该对象代表当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的文本字符区域。只读。

## 说明

**[Characters](./../../Characters/Characters 对象.html)** 对象不是集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Start | any | 可选 | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。 |
| Length | any | 可选 | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（*Start* 字符之后的所有字符）。 |

## 返回值

Characters

## 示例

```javascript
/*此示例显示 A3 单元格文本中前两个字符的字体颜色是否为红色。*/
function test() {
    console.log(Range("A3").DisplayFormat.Characters(1, 2).Font.Color == RGB(255, 0, 0))
}
```

```javascript
/*此示例显示第一个工作表中 B5 单元格文本的字体大小。*/
function test() {
    console.log(Worksheets.Item(1).Range("B5").DisplayFormat.Characters().Font.Size)
}
```
