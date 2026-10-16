---
title: "Font 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UniqueValues/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
---
# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，该对象在条件格式规则的计算结果为 **True** 时指定字体格式。只读。

## 说明

条件格式对象并不支持 **[Font](./../../Font/Font 对象.html)** 对象的所有属性。您可以设置字体样式、下划线、颜色和删除线属性。

## 示例

```javascript
/*本示例设置活动工作表上区域 C1:C10 中第二个（UniqueValues）条件格式的特定文本的字体的颜色和字号。*/
function test() {
    let uniqueValues = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    uniqueValues.Font.ColorIndex = 4
    uniqueValues.Font.Size = 15
}
```

```javascript
/*本示例将活动工作表上区域 C1:C10 中第一个（UniqueValues）条件格式特定文本的字体设置为加粗和倾斜，并设置该字体的颜色。*/
function test() {
    let uniqueValues = ActiveSheet.Range("C1:C10").FormatConditions.Item(1)
    uniqueValues.Font.Bold = true
    uniqueValues.Font.Italic = true
    uniqueValues.Font.Color = RGB(255, 0, 0)
}
```
