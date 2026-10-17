---
title: "AppliesTo (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UniqueValues/%E5%B1%9E%E6%80%A7/AppliesTo%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D-9T0Poi
---

# AppliesTo (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，指定格式规则将应用于的单元格区域。

## 示例

```javascript
/*本示例选中工作表 Sheet1 上区域 C1:C10 中第一个（UniqueValues）条件格式应用到的单元格区域。*/
function test() {
    let uniqueValues = Application.Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    uniqueValues.AppliesTo.Select()
}
```

```javascript
/*本示例将活动工作表上区域 C1:C10 中第二个（UniqueValues）条件格式应用到的单元格区域的地址赋值给 A1 单元格。*/
function test() {
    let uniqueValues = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    Range("A1").Value2 = uniqueValues.AppliesTo.Address()
}
```
