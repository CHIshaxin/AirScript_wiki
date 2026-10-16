---
title: "TintAndShade 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/TintAndShade%20%E5%B1%9E%E6%80%A7.html
---
# TintAndShade (属性)

返回或设置一个 **Single**，使颜色变深或变浅。

## 说明

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

## 示例

```javascript
/*本示例将活动工作表 A1:B5 区域单元格内部设置为红色，并将亮度设置为0.5。*/
function test() {
    let interior = ActiveSheet.Range("A1:B5").Interior
    interior.ColorIndex = 3
    interior.TintAndShade = 0.5
}
```

```javascript
/*本示例显示 Sheet1 中 A4 单元格内部背景的亮度值。*/
function test() {
    console.log(Application.Worksheets.Item("Sheet1").Range("A4").Interior.TintAndShade)
}
```
