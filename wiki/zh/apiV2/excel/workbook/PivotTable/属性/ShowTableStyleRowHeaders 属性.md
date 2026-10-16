---
title: "ShowTableStyleRowHeaders (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ShowTableStyleRowHeaders%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dgnatitp
---

# ShowTableStyleRowHeaders (属性)

如果行标题应显示在数据透视表中，则将 **ShowTableStyleRowHeaders** 属性设置为 **True**。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例将行标题显示在数据透视表上*/
function test() {
    Worksheets.Item(1).PivotTables(1).ShowTableStyleRowHeaders = true
}
```

```javascript
/*本示例显示数据透视表的 ShowTableStyleRowHeaders 属性的数据类型，并设置为 false*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).ShowTableStyleRowHeaders)
    ActiveSheet.PivotTables(1).ShowTableStyleRowHeaders = false
}
```
