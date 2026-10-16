---
title: "UseStandardHeight (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/UseStandardHeight%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DrrRoVUE
---

# UseStandardHeight (属性)

如果 **[Range](./../../Range/Range 对象.html)** 对象的行高等于工作表的标准行高，则该值为 **True**。如果区域包含不止一行并且不是所有的行都等高，则返回 **Null**。**Variant** 类型，可读写。

## 示例

```javascript
/*本示例将 Sheet1 中的首行的行高设置为标准行高。*/
function test() {
    Worksheets.Item("Sheet1").Rows.Item(1).UseStandardHeight = true
}
```

```javascript
/*本示例将 A1 单元格所在行的行高设置为 50，并显示 A1 单元格和 A1:A5 单元格区域的行高是否等于标准行高*/
function test() {
    Range("A1").Rows.RowHeight = 50
    console.log(Range("A1").UseStandardHeight) // false
    console.log(Range("A1:A5").UseStandardHeight) // null
}
```
