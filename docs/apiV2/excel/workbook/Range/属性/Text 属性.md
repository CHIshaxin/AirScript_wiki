---
title: "Text 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Text%20%E5%B1%9E%E6%80%A7.html
---
# Text (属性)

返回或设置指定对象中的文本。**String** 型，只读。

## 示例

```javascript
/*此示例演示包含格式数字的单元格的 Text 和 Value 属性的区别。*/
function test() {
    let c = Application.Worksheets.Item("Sheet1").Range("B14")
    c.Value2 = 1198.3
    c.NumberFormat = "$#,##0_);($#,##0)"
    console.log(c.Value2)
    console.log(c.Text)
}
```

```javascript
/*本示例将 A1 单元格的值设置为 44958，设置数值格式为日期，并显示 A1 单元格的文本*/
function test() {
    Range("A1").Value2 = 44958
    Range("A1").NumberFormat = "yyyy/mm/dd"
    console.log(Range("A1").Text)
}
```
