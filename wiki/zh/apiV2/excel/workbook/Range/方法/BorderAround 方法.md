---
title: "BorderAround (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/BorderAround%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: De0Ebk_i
---

# BorderAround (方法)

向单元格区域添加边框，并设置该新边框的 **Color**、**LineStyle** 和 **Weight** 属性。**Variant** 类型。

## 说明

必须只指定以下一项：*ColorIndex*、*Color* 或 *ThemeColor*。

可指定 *LineStyle* 或 *Weight*，但不能同时指定这两者。如果未指定任何参数，ET 将使用默认的线条样式和粗细。

本方法只绘制了整个区域的外框，而不填充该区域。如果要设置所有单元格的边框，就必须设置 **[Borders](./../../Borders/Borders 对象.html)** 集合的 **Color**、**LineStyle** 和 **Weight** 属性。要清除边框，必须将区域中所有单元格的 **LineStyle** 属性设置为 **xlLineStyleNone**。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| LineStyle | any | 可选 | 指定边框线型的 **[XlLineStyle](./../../枚举/XlLineStyle 枚举.html)** 常量之一。 |
| Weight | Default | 可选 | 边框粗细。 |
| Color | any | 可选 | 边框颜色，以 RGB 值表示。 |
| ThemeColor | any | 可选 | 主题颜色，作为当前颜色主题的索引，或作为 **[XlThemeColor](./../../枚举/XlThemeColor 枚举.html)** 值。 |

## 返回值

Variant

## 示例

```javascript
/*本示例在 Sheet1 的区域 A1:D4 周围添加红色粗边框。*/
function test() {
    Worksheets("Sheet1").Range("A1:D4").BorderAround(undefined, xlThick, 3)
}
```

```javascript
/*本示例将 A1:C3 单元格区域边框设置为实线、中等宽度、通过 RGB 设置红色*/
function test() {
    Range("A1:C3").BorderAround(xlContinuous, xlMedium, null, RGB(255, 0, 0))
}
```
