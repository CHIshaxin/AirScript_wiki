---
title: "AppliesTo 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/AppliesTo%20%E5%B1%9E%E6%80%A7.html
---
# AppliesTo (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，指定格式规则将应用于的单元格区域。

## 示例

```javascript
/*本示例选中工作表 Sheet1 上区域 A1:A10 中第一个为数据条的条件格式应用到的单元格区域。*/
function test() {
    let databar = Application.Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(1)
    databar.AppliesTo.Select()
}
```

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个为数据条的条件格式应用到的单元格区域的地址。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.AppliesTo.Address())
}
```
