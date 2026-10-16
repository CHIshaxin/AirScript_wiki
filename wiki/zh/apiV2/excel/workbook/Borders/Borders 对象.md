---
title: "Borders (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Borders/Borders%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Bh4F652f
---

# Borders (对象)

由四个 **[Border](./../Border/Border 对象.html)** 对象组成的集合，它们分别代表 **Range**或 **[Style](./../Style/Style 对象.html)** 对象的四个边框。

## 说明

使用 **Borders** 属性可返回包含所有四个边框的 **[Borders](./../Borders/Borders 对象.html)** 集合。 可以对单元格或区域的每一侧应用不同的边框。 有关如何对单元格区域应用边框的详细信息，请参阅 **Range.Borders** 属性。

只能对 **Range** 和 **[Style](./../Style/Style 对象.html)** 对象的各个边框分别设置边框属性。 其他带边框的对象（如误差线和序列线）具有被视为单个实体的边框，而不管它有多少边。 对于这些对象，在返回边框和设置边框属性时必须将其作为一个整体处理。 有关详细信息，请参阅 **[Border](./../Border/Border 对象.html)** 对象。

## 示例

```javascript
/*本示例向第一张工作表上的单元格 A1 添加双边框。*/
function test() {
    Worksheets.Item(1).Range("A1").Borders.LineStyle = xlDouble
}
```

使用 **Borders** (索引) （其中 index 标识边框）返回单个 **[Border](./../Border/Border 对象.html)** 对象。 Index 可以是以下 **[XlBordersIndex](./../枚举/XlBordersIndex 枚举.html)** 常量之一： **xlDiagonalDown** 、 **xlDiagonalUp** 、 **xlEdgeBottom** 、 **xlEdgeLeft** 、 **xlEdgeRight** 、 **xlEdgeTop** 、 **xlInsideHorizontal** 或 **xlInsideVertical** 。

```javascript
/*本示例将工作表 Sheet1 中单元格区域 A1:G1 的底部边框颜色设置为红色。*/
function test() {
    Worksheets("Sheet1").Range("A1:G1").Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
}
```

```javascript
/*本示例在区域中的所有单元格周围生成一个细边框。*/
function test() {
    let borders = Range("B6:D8").Borders
    borders.LineStyle = xlContinuous
    borders.Weight = xlThin
}
```

```javascript
/*本示例更改区域的内部单元格边框。*/
function test() {
    let borders = Range("B2:D4").Borders
    borders.Item(xlInsideHorizontal).LineStyle = xlContinuous
    borders.Item(xlInsideHorizontal).Weight = xlThin
    borders.Item(xlInsideVertical).LineStyle = xlContinuous
    borders.Item(xlInsideVertical).Weight = xlThin
}
```
