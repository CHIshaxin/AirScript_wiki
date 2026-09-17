---
title: "TintAndShade 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatColor/%E5%B1%9E%E6%80%A7/TintAndShade%20%E5%B1%9E%E6%80%A7.html
---
# TintAndShade (属性)

返回或设置一个 **Single** 类型的值，该值使用于数据条或色阶条件格式规则的单元格的填充颜色变浅或变深。

## 说明

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

## 示例

```javascript
/*本示例显示指定单元格区域的第一个色阶条件格式中第一个条件的最小阈值的填充颜色是否为中间值。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.ColorScaleCriteria.Item(1).FormatColor.TintAndShade == 0)
}
```

```javascript
/*本示例选中指定单元格区域，设置第一个色阶条件格式中第二个条件，将最大阈值的填充颜色设置为最暗。*/
function test() {
    Range("C1:C10").Select()
    Selection.FormatConditions.Item(1).ColorScaleCriteria.Item(2).FormatColor.TintAndShade = -1
}
```
