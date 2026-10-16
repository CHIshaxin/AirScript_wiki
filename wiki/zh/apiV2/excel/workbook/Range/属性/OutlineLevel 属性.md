---
title: "OutlineLevel (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/OutlineLevel%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cq3WFKAG
---

# OutlineLevel (属性)

返回或设置指定行或列的当前分级显示级别。**Variant** 类型，可读写。

## 说明

第一级为最外层的摘要级。

## 示例

```javascript
/*本示例设置 sheet1 上第二行的分级显示级别。*/
function test() {
    Worksheets.Item("Sheet1").Rows.Item(2).OutlineLevel = 1
}
```

```javascript
/*本示例将 A3 单元格所在行的分级显示级别设置为2，并显示*/
function test() {
    Range("A3").Rows.OutlineLevel = 2
    console.log(Range("A3").Rows.OutlineLevel)
}
```
