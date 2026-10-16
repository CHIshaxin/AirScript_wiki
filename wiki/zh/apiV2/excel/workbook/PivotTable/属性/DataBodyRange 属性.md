---
title: "DataBodyRange (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/DataBodyRange%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C7XVRGoz
---

# DataBodyRange (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表数据透视表中数值区域。只读。

## 示例

```javascript
/*本示例将活动工作表中第一张数据透视表数值区域的第一个单元格内容改为KingSoft*/
function test() {
    ActiveSheet.PivotTables(1).DataBodyRange.Item(1).Value2 = "KingSoft"
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表数值区域的地址*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables(1).DataBodyRange.Address())
}
```
