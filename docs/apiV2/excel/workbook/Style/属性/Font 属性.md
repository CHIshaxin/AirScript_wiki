---
title: "Font 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
---
# Font (属性)

返回一个Font 对象，它代表指定对象的字体。

## 示例

```javascript
/*本示例将第一张工作表中A1单元格字体的字号设为12磅。*/
function test() {
    let range = Application.Worksheets.Item(1).Range("A1")
    range.Value2 = "这是示例文本"
    range.Style.Font.Size = 12
}
```

```javascript
/*本示例将活动工作表中B1单元格的字体设置为绿色。*/
function test() {
    let range = ActiveSheet.Range("B1")
    range.Value2 = "这是示例文本"
    range.Style.Font.Color = RGB(0, 255, 0)
}
```
