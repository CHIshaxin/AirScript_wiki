---
title: "Color (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatColor/%E5%B1%9E%E6%80%A7/Color%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C2vXVPGR
---

# Color (属性)

返回或设置与数据条或色阶条件格式规则的阈值关联的填充色。

## 说明

格式颜色以 RGB 函数表示。例如，要将颜色设置为红色，请使用 `RGB(255,0,0)`。

## 示例

```javascript
/*本示例选中指定单元格区域，设置第一个色阶条件格式中第一个条件，将最小阈值的颜色改为绿色。*/
function test() {
    Range("C1:C10").Select()
    Selection.FormatConditions.Item(1).ColorScaleCriteria.Item(1).FormatColor.Color = RGB(0, 255, 0)
}
```

```javascript
/*本示例显示指定单元格区域的第一个色阶条件格式中第二个条件的最大阈值颜色是否为红色。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.ColorScaleCriteria.Item(2).FormatColor.Color == RGB(255, 0, 0))
}
```
