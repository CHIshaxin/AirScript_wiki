---
title: "Locked 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
---
# Locked (属性)

返回或设置一个 **Variant** 值，它指明对象是否已被锁定。

## 说明

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**；如果指定区域既包含锁定单元格又包含不锁定单元格，则返回 **Null**。

## 示例

```javascript
/*此示例解除对 Sheet1 中 A1:G37 区域单元格的锁定，以便当该工作表受保护时也可对这些单元格进行修改。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:G37").Locked = false
    Worksheets.Item("Sheet1").Protect()
}
```

```javascript
/*本示例检查 A4:B6 单元格区域是否被锁定，然后将 A4 单元格解除锁定，再显示区域的锁定情况*/
function test() {
    console.log(Range("A4:B6").Locked) //true
    Range("A4").Locked = false
    console.log(Range("A4:B6").Locked) //null
}
```
