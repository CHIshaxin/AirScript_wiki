---
title: "TableRange1 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/TableRange1%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D9itFFT1
---

# TableRange1 (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示包括整个数据透视表但不包括页字段的区域。只读。

## 说明

**TableRange2** 属性包括页字段。

## 示例

```javascript
/*本示例选定整个数据透视表，但页字段除外。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    Range("F3").PivotTable.TableRange1.Select()
}
```

```javascript
/*本示例显示数据透视表的地址，不包括页字段*/
function test() {
    console.log(ActiveSheet.PivotTables(1).TableRange1.Address())
}
```
