---
title: "ShowDrillIndicators 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ShowDrillIndicators%20%E5%B1%9E%E6%80%A7.html
---
# ShowDrillIndicators (属性)

**ShowDrillIndicators** 属性用于切换数据透视表中深化指示符的显示。可读/写 **Boolean** 类型。

## 说明

如果深化指示符显示在数据透视表中，则返回 **True**；如果深化指示符未显示在数据透视表中，则返回 **False**。

## 示例

```javascript
/*本示例显示数据透视表中的深化指示符*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).ShowDrillIndicators = true
}
```

```javascript
/*本示例显示第一张数据透视表 ShowDrillIndicators 属性的数据类型，并设置为 false。*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).ShowDrillIndicators)
    ActiveSheet.PivotTables(1).ShowDrillIndicators = false
}
```
