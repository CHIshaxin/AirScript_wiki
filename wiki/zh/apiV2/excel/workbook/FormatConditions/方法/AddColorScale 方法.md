---
title: "AddColorScale (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/AddColorScale%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: NKKVzfHW
---

# AddColorScale (方法)

返回一个代表条件格式规则的新 **[ColorScale](./../../ColorScale/ColorScale 对象.html)** 对象，该条件格式规则使用单元格颜色渐变来表示选定区域中所含单元格值的相对差异。

## 说明

色阶是直观的参照，可以帮助您了解数据的分布和变化。色阶可帮助您利用颜色的变化来标出给定区域中单元格值的相对差异。不同的颜色和颜色之间的渐变代表单元格值的差异。例如，在三色色阶中，您可以指定包含最高相对数据值的单元格为绿色，包含中间值的单元格为黄色，包含最低值的单元格为红色。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| ColorScaleType | int4 | 必填 | 色阶的类型。 |

## 示例

```javascript
/*本示例在活动工作表上单元格区域 C1:C10 新增色阶条件格式。*/
function test() {
    ActiveSheet.Range("C1:C10").FormatConditions.AddColorScale(3)
}
```

```javascript
/*本示例在活动工作表上单元格区域 A1:A10 新增色阶条件格式。*/
function test() {
    let colorScale = ActiveSheet.Range("A1:A10").FormatConditions.AddColorScale(2)
    colorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria(2).FormatColor.ColorIndex = 7
}
```
