---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScaleCriteria/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从 **[ColorScaleCriteria](./../../ColorScaleCriteria/ColorScaleCriteria 对象.html)** 集合中返回单个 **ColorScaleCriterion** 对象。只读。

## 说明

*Index* 参数的值不能大于为色阶条件格式设置的条件的数目。条件与为色阶分配的阈值等效。要查找阈值的数目，请使用 **ColorScaleCriteria.Count** 属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | **ColorScaleCriterion** 对象的索引号。 |

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
/*本示例判断如果活动工作表上单元格区域 A1:A10 的第一个色阶条件格式的第二个阈值条件的类型为xlConditionValuePercentile，则设置该阈值条件的主题颜色和亮度。*/
function test() {
    let colorScale = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    let criterion = colorScale.ColorScaleCriteria.Item(2)
    if (criterion.Type == xlConditionValuePercentile) {
        criterion.FormatColor.ThemeColor = xlThemeColorAccent6
        criterion.FormatColor.TintAndShade = 0.5
    }
}
```
