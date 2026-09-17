---
title: "Style 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/Style%20%E5%AF%B9%E8%B1%A1.html
---
# Style (对象)

代表区域的样式说明。

## 说明

**[Style](./../Style/Style 对象.html)** 对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置样式，包括“常规”、“货币”和“百分比”。同时对多个单元格修改单元格格式属性时，使用 **[Style](./../Style/Style 对象.html)** 对象是快捷高效的方法。

对于 **[Workbook](./../Workbook/Workbook 对象.html)** 对象， **[Style](./../Style/Style 对象.html)** 对象是 **[Styles](./../Styles/Styles 对象.html)** 集合的成员。 **[Styles](./../Styles/Styles 对象.html)** 集合包含该工作簿的所有已定义样式。

通过更改应用于单元格的样式的属性可更改单元格的外观。但要记住，更改样式的属性将影响所有以该样式格式化了的单元格。

样式按照名称的字母顺序排序。样式编号表明指定样式在样式名排序列表中的位置。`Styles(1)` 是排序列表中的第一个样式，而 `Styles(Styles.Count)` 是最后一个。

有关创建和修改样式的详细信息，请参阅 **[Styles](./../Styles/Styles 对象.html)** 对象。

使用 **Style** 属性可返回一个用于 **[Range](./../Range/Range 对象.html)** 对象的 **[Style](./../Style/Style 对象.html)** 对象。

## 示例

```javascript
/*本示例将“百分比”样式应用于 Sheet1 中的单元格区域 A1:A10。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:A10").Style = "Percent"
}
```

使用 **Styles**(*index*)（其中 *index* 是样式索引号或名称）可从工作簿的 **[Style](./../Style/Style 对象.html)** 集合中返回一个 **[Styles](./../Styles/Styles 对象.html)** 对象。

```javascript
/*本示例通过设置样式的 Bold 属性来更改活动工作簿的 Normal 样式。*/
function test() {
    Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
}
```
