---
title: "Characters 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E6%96%B9%E6%B3%95/Characters%20%E6%96%B9%E6%B3%95.html
---
# Characters (方法)

返回 **[Characters](./../../Characters/Characters 对象.html)** 对象，它代表对象文本内某个区域的字符。使用 **[Characters](./../../Characters/Characters 对象.html)** 对象可为文本字符串内的字符设置格式。

## 说明

**[Characters](./../../Characters/Characters 对象.html)** 对象不是集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Start | any | 可选 | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。 |
| Length | any | 可选 | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（Start 字符之后的所有字符）。 |

## 示例

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第一个数据标签设置为倾斜。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1)
    datalabel.Characters().Font.Italic = true
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第三个数据标签的文本设置为“五年级数学成绩”，并将该数据标签的前三个字符设置为加粗。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(3)
    datalabel.Caption = "五年级数学成绩"
    datalabel.Characters(null, 3).Font.Bold = true
}
```
