---
title: "Worksheet (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Worksheet%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C58ekLiX
---

# Worksheet (属性)

返回一个 **[Worksheet](./../../Worksheet/Worksheet 对象.html)** 对象，该对象表示包含指定区域的工作表。只读。

## 示例

```javascript
/*本示例显示包含活动单元格的工作表的名称。本示例必须在工作表上运行。*/
function test() {
    console.log(Application.ActiveCell.Worksheet.Name)
}
```

```javascript
/*本示例显示包含名为“testRange”的区域的工作表名称。*/
function test() {
    console.log(Application.Range("testRange").Worksheet.Name)
}
```
