---
title: "HasFormula (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/HasFormula%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B-yIesAm
---

# HasFormula (属性)

如果区域中所有单元格均包含公式，则该属性值为 **True**；如果所有单元格均不包含公式，则该属性值为 **False**；其他情况下为 **null**。**Variant** 类型，只读。

## 示例

```javascript
/*本示例提示用户选定 sheet1 的区域。如果该选定区域内所有单元格均包含公式，则本示例将显示消息框。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    let rr = Application.InputBox("Select a range on this worksheet", undefined, undefined, undefined, undefined, undefined, undefined, 8)
    if (rr.HasFormula) {
        console.log("Every cell in the selection contains a formula")
    }
}
```

```javascript
/*本示例检查 B1:C1 单元格区域是否所有单元格都含有公式，然后给他们分别设置公式，查看结果*/
function test() {
    console.log(Range("B1:C1").HasFormula) //false
    Range("B1").Formula = "=B2+B3"
    console.log(Range("B1:C1").HasFormula) //null
    Range("C1").Formula = "=C2+C3"
    console.log(Range("B1:C1").HasFormula) //true
}
```
