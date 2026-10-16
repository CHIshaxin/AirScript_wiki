---
title: "TableRange2 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/TableRange2%20%E5%B1%9E%E6%80%A7.html
---
# TableRange2 (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示包括整个数据透视表（含页字段）的区域。只读。

## 说明

**TableRange1**属性不包括页字段。

## 示例

```javascript
/*本示例选定整个数据透视表，包括页字段。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    Range("F3").PivotTable.TableRange2.Select()
}
```

```javascript
/*本示例显示数据透视表的地址，包括页字段。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).TableRange2.Address())
}
```
