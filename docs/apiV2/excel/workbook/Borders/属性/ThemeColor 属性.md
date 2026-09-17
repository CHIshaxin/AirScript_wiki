---
title: "ThemeColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Borders/%E5%B1%9E%E6%80%A7/ThemeColor%20%E5%B1%9E%E6%80%A7.html
---
# ThemeColor (属性)

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

## 说明

如果对象当前未应用主题颜色，试图访问其主题颜色则会引起无效请求运行时错误。

## 示例

```javascript
/*本示例将 A2 单元格的主题颜色设置为指定“强调文字颜色6”主题颜色。*/
function test() {
    Range("A2").Borders.ThemeColor = msoThemeColorAccent6
}
```

```javascript
/*本示例显示 E10 单元格的主题颜色是否为指定无主题颜色。*/
function test() {
    console.log(Range("E10").Borders.ThemeColor == msoNotThemeColor)
}
```
