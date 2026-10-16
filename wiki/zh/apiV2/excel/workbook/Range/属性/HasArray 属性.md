---
title: "HasArray (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/HasArray%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BwOCO7gh
---

# HasArray (属性)

如果指定单元格属于数组公式，则该属性值为 **True**。**Variant** 类型，只读。

## 示例

```javascript
/*如果 sheet1 上的活动单元格为数组的一部分，则本示例将显示一条消息。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    if (ActiveCell.HasArray) {
        console.log("The active cell is part of an array")
    }
}
```

```javascript
/*本示例将 A1:C4 单元格区域设置数组公式 {=3}，显示 B3 单元格是否属于数组公式*/
function test() {
    Range("A1:C4").FormulaArray = "=3"
    console.log(Range("B3").HasArray)
}
```
