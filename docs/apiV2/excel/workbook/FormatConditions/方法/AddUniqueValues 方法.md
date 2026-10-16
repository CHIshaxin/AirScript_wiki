---
title: "AddUniqueValues 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/AddUniqueValues%20%E6%96%B9%E6%B3%95.html
---
# AddUniqueValues (方法)

返回一个代表指定区域条件格式规则的新 **[UniqueValues](./../../UniqueValues/UniqueValues 对象.html)** 对象。

## 说明

可以使用 **[UniqueValues](./../../UniqueValues/UniqueValues 对象.html)** 对象快速显现包含唯一值或重复值的单元格。

## 示例

```javascript
/*本示例在活动工作表上单元格区域 A1:A10 新增UniqueValues条件格式。*/
function test() {
    let uniqueValues = ActiveSheet.Range("A1:A10").FormatConditions.AddUniqueValues()
    uniqueValues.Font.Bold = true
    uniqueValues.Font.ColorIndex = 3
}
```

```javascript
/*本示例在活动工作表上单元格区域 B1:B10 上新增UniqueValues条件格式。*/
function test() {
    let uniqueValues = ActiveSheet.Range("B1:B10").FormatConditions.AddUniqueValues()
    uniqueValues.Interior.Color = RGB(0, 255, 255)
}
```
