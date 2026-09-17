---
title: "TintAndShade 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/TintAndShade%20%E5%B1%9E%E6%80%A7.html
---
# TintAndShade (属性)

返回或设置一个 **Single**，使颜色变深或变浅。

## 说明

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

## 示例

```javascript
/*本示例将 A2 单元格的样式中字体的 TintAndShade 属性设置为 1。*/
function test() {
    Range("A2").Style.Font.TintAndShade = 1
}
```

```javascript
/*本示例显示活动工作簿第一个样式的字体的 TintAndShade 属性值。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).Font.TintAndShade)
}
```
