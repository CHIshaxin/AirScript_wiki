---
title: "Styles (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Styles/Styles%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: y-n0Fe0B
---

# Styles (对象)

指定工作簿或活动工作簿中所有 **[Style](./../Style/Style 对象.html)** 对象的集合。

## 说明

每一个 **[Style](./../Style/Style 对象.html)** 对象都代表对某区域的样式描述。 **[Style](./../Style/Style 对象.html)** 对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置的样式，包括“常规”、“货币”和“百分比”。

使用 **Styles** 属性可返回 **[Styles](./../Styles/Styles 对象.html)** 集合。

## 示例

```javascript
/*此示例在第一个工作表上创建活动工作簿中样式名的列表。*/
function test() {
    for (let i = 1; i <= Application.ActiveWorkbook.Styles.Count; i++) {
        Application.Worksheets.Item(1).Cells.Item(i, 1).Value2 = ActiveWorkbook.Styles.Item(i).Name
    }
}
```

使用 **Add** 方法可创建一个新的样式并将它添加到集合。

```javascript
/*此示例基于“常规”样式创建一个新的样式，修改边框和字体，然后将该新样式应用到单元格 A25:A30。*/
function test() {
    let style = Application.ActiveWorkbook.Styles.Add("Bookman Top Border")
    style.Borders.Item(xlEdgeTop).LineStyle = xlDouble
    style.Font.Bold = true
    style.Font.Name = "Bookman"
    Application.Worksheets.Item(1).Range("A25:A30").Style = "Bookman Top Border"
}
```

使用 **Styles**(*index*)（其中 *index* 是样式索引号或名称）可从工作簿的 **[Style](./../Style/Style 对象.html)** 集合中返回一个 **[Styles](./../Styles/Styles 对象.html)** 对象。

```javascript
/*此示例通过设置活动工作簿中“常规”样式的 Bold 属性来更改该样式。*/
function test() {
    Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
}
```
