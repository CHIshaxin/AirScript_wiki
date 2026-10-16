---
title: "ShowTableStyleLastColumn (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ShowTableStyleLastColumn%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DYhvj9eS
---

# ShowTableStyleLastColumn (属性)

如果为指定的 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象显示最后一列，则返回或设置。读/写 **Boolean**。

## 示例

```javascript
/*本示例取消显示数据透视表的最后一列*/
function test() {
    Worksheets.Item(1).PivotTables(1).ShowTableStyleLastColumn = false
}
```

```javascript
/*本示例显示数据透视表 ShowTableStyleLastColumn 属性的数据类型，并设置为 true*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).ShowTableStyleLastColumn)
    ActiveSheet.PivotTables(1).ShowTableStyleLastColumn = true
}
```
