---
title: "HasAutoFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/HasAutoFormat%20%E5%B1%9E%E6%80%A7.html
---
# HasAutoFormat (属性)

如果在刷新数据透视表或移动字段时自动设置格式，则该属性值为 **True**。 读/写 **Boolean**。

## 示例

```javascript
/*本示例设置刷新数据透视表或移动字段时自动设置格式*/
function test() {
    Worksheets.Item(1).PivotTables(1).HasAutoFormat = true
}
```

```javascript
/*本示例显示数据透视表 HasAutoFormat 属性的数据类型，并设置为 false。*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).HasAutoFormat)
    ActiveSheet.PivotTables(1).HasAutoFormat = false
}
```
