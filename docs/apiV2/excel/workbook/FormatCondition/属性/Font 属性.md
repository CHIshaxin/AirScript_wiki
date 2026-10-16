---
title: "Font 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
---
# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，它代表指定对象的字体。

## 示例

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式的特定文本的字体的颜色和字号。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.Font.ColorIndex = 4
    formatCondition.Font.Size = 15
}
```

```javascript
/*本示例将活动工作表上区域 A1:A10 中第一个条件格式特定文本的字体设置为加粗和倾斜。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.Font.Bold = true
    formatCondition.Font.Italic = true
}
```
