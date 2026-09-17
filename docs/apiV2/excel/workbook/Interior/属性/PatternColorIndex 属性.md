---
title: "PatternColorIndex 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/PatternColorIndex%20%E5%B1%9E%E6%80%A7.html
---
# PatternColorIndex (属性)

返回或设置内部图案的颜色，表示为当前调色板中的颜色编号或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。**Long** 类型，可读写。

## 说明

将本属性设置为 **xlColorIndexAutomatic**，表示单元格使用自动图案或图形对象使用自动填充样式。将本属性设置为 **xlColorIndexNone**，则表示不希望使用图案（等价于将 **[Interior](./../../Interior/Interior 对象.html)** 对象的 **Pattern** 属性设置为 **xlPatternNone**）。

## 示例

```javascript
/*本示例对 Sheet1 中 A1:B2 区域单元格添加棋盘图案，并将图案设置为蓝色。*/
function test() {
    let interior = Worksheets.Item("Sheet1").Range("A1:B2").Interior
    interior.Pattern = xlChecker
    interior.PatternColorIndex = 5
}
```

```javascript
/*本示例对活动工作表中 D4 单元格添加网格图案，并将图案颜色设置为自动配色。*/
function test() {
    let interior = ActiveSheet.Range("D4").Interior
    interior.Pattern = xlPatternGrid
    interior.PatternColorIndex = xlColorIndexAutomatic
}
```
