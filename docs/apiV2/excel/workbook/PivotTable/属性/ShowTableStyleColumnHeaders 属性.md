---
title: "ShowTableStyleColumnHeaders 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ShowTableStyleColumnHeaders%20%E5%B1%9E%E6%80%A7.html
---
# ShowTableStyleColumnHeaders (属性)

如果列标题应显示在数据透视表中，则将 **ShowTableStyleColumnHeaders** 属性设置为 **True**。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例取消列标题在数据透视表中的显示*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).ShowTableStyleColumnHeaders = false
}
```

```javascript
/*本示例显示第一张数据透视表 ShowTableStyleColumnHeaders 属性的数据类型，并设置为 true。*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).ShowTableStyleColumnHeaders)
    ActiveSheet.PivotTables(1).ShowTableStyleColumnHeaders = true
}
```
