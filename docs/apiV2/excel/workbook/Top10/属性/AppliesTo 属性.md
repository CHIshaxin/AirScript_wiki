---
title: "AppliesTo 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Top10/%E5%B1%9E%E6%80%A7/AppliesTo%20%E5%B1%9E%E6%80%A7.html
---
# AppliesTo (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，指定格式规则将应用于的单元格区域。

## 示例

```javascript
/*本示例选中工作表 Sheet1 上区域 B1:B10 中第一个（Top10）条件格式应用到的单元格区域。*/
function test() {
    let top = Application.Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions.Item(1)
    top.AppliesTo.Select()
}
```

```javascript
/*本示例将活动工作表上区域 B1:B10 中第二个（Top10）条件格式应用到的单元格区域的地址赋值给 C1 单元格。*/
function test() {
    let top = ActiveSheet.Range("B1:B10").FormatConditions.Item(2)
    Range("C1").Value2 = top.AppliesTo.Address()
}
```
