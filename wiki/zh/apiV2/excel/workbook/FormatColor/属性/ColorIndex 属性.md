---
title: "ColorIndex (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatColor/%E5%B1%9E%E6%80%A7/ColorIndex%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DlJ1PLoM
---

# ColorIndex (属性)

返回或设置 **XlColorIndex**枚举的常量之一，指定是否以当前调色板中颜色的索引值形式表示填充色。

## 说明

此属性用于色阶或数据条条件格式规则中的每个阈值。

## 示例

```javascript
/*本示例选中指定单元格区域，设置第二个色阶条件格式中第一个条件，将最小阈值的颜色索引设置为自动配色。*/
function test() {
    Range("C1:C10").Select()
    Selection.FormatConditions.Item(2).ColorScaleCriteria.Item(1).FormatColor.ColorIndex = xlColorIndexAutomatic
}
```

```javascript
/*本示例显示指定单元格区域的第一个色阶条件格式中第二个条件最大阈值的颜色索引是否为无颜色。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.ColorScaleCriteria.Item(2).FormatColor.ColorIndex == xlColorIndexNone)
}
```
