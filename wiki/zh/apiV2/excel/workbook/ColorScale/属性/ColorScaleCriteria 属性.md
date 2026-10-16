---
title: "ColorScaleCriteria (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/%E5%B1%9E%E6%80%A7/ColorScaleCriteria%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BsQJQOpw
---

# ColorScaleCriteria (属性)

返回一个 **[ColorScaleCriteria](./../../ColorScaleCriteria/ColorScaleCriteria 对象.html)** 对象，该对象是单个 **ColorScaleCriterion** 对象的集合。**ColorScaleCriterion** 对象指定用于色阶条件格式的阈值条件的类型、值和颜色。只读。

## 示例

```javascript
/*本示例设置活动工作表上单元格区域 A1:A10 的第二个色阶条件格式的阈值条件的颜色。*/
function test() {
    let colorScale = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    colorScale.ColorScaleCriteria.Item(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria.Item(2).FormatColor.ColorIndex = 7
}
```

```javascript
/*本示例判断如果活动工作表上单元格区域 A1:A10 的第一个色阶条件格式的第三个阈值条件的类型为xlConditionValueHighestValue，则设置该阈值条件的主题颜色和亮度。*/
function test() {
    let colorScale = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    let criterion = colorScale.ColorScaleCriteria.Item(3)
    if (criterion.Type == xlConditionValueHighestValue) {
        criterion.FormatColor.ThemeColor = xlThemeColorAccent3
        criterion.FormatColor.TintAndShade = 0.5
    }
}
```
