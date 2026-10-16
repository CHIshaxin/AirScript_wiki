---
title: "Hidden (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Hidden%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: tPz3qb0J
---

# Hidden (属性)

返回或设置一个 **Variant** 值，它指明是否隐藏行或列。

## 说明

将此属性设置为 **True** 以隐藏行或列。指定的区域必须占据整个行或整个列。

请不要将此属性与 **FormulaHidden** 属性混淆。

## 示例

```javascript
/*此示例隐藏 Sheet1 的 C 列。*/
function test() {
    Application.Worksheets.Item("Sheet1").Columns.Item("C").Hidden = true
}
```

```javascript
/*本示例隐藏 A1 单元格所在的行*/
function test() {
    Range("A1").Rows.Hidden = true
}
```
