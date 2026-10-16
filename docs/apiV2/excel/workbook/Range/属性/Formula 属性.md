---
title: "Formula 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Formula%20%E5%B1%9E%E6%80%A7.html
---
# Formula (属性)

返回或设置一个 **Variant** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

## 说明

此属性对于 OLAP 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，**Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 JavaScript 中相同维数的数组。同样，也可在 JavaScript 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

## 示例

```javascript
/*此示例设置 Sheet1 中 A1 单元格的公式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=$A$4+$A$10"
}
```

```javascript
/*本示例将 A2:A3 单元格区域设置日期格式：1988年5月*/
function test() {
    Range("A2:A3").Formula = "1988年5月"
}
```
