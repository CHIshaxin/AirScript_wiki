---
title: "NoteText 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/NoteText%20%E6%96%B9%E6%B3%95.html
---
# NoteText (方法)

返回或设置与区域左上角单元格相关联的单元格注释。**String** 类型，可读写。单元格注释已经替换为区域注释。有关详细信息，请参阅 **[Comment](./../../Comment/Comment 对象.html)** 对象。

## 说明

若要添加超过 255 个字符的注释，请先使用此方法指定前 255 个字符，然后再用它为注释追加其余部分（一次不超过 255 个字符）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Text | any | 可选 | 要添加到注释中的文本（不超过 255 个字符）。注释文本插入的起始位置由 *Start* 参数指定，并替换现有注释的 *Length* 字符。如果省略该参数，此方法将返回注释的当前文本，该注释起始于 *Start* 位置，长度为 *Length* 个字符。 |
| Start | any | 可选 | 设置或返回的文本的起始位置。如果省略该参数，此方法将第一个字符作为起始位置。要将文本追加到注释，请指定一个大于现有注释中的字符数的数字。 |
| Length | any | 可选 | 要设置或返回的字符数。如果省略该参数，则 ET 将以从现有注释起始位置到注释末尾的长度设置或返回字符（不超过 255 个字符）。如果从 *Start* 位置到注释的末尾的字符数超过 255，则此方法只返回 255 个字符。 |

## 返回值

String

## 示例

```javascript
/*本示例为工作表 Sheet1 上的单元格 A1 设置单元格注释。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").NoteText("This may change!")
}
```

```javascript
/*本示例给 C2 单元格添加批注 This is C2，将第 9 个之后的字符替换为 test*/
function test() {
    Range("C2").AddComment("This is C2")
    Range("C2").NoteText("test", 9)
}
```
