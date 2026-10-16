---
title: "AppliesTo (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/AppliesTo%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C5-AyNTH
---

# AppliesTo (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，指定格式规则将应用于的单元格区域。

## 示例

```javascript
/*本示例选中工作表 Sheet1 上区域 B1:B10 中第一个条件格式应用到的单元格区域。*/
function test() {
    let formatCondition = Application.Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions.Item(1)
    formatCondition.AppliesTo.Select()
}
```

```javascript
/*本示例显示活动工作表上区域 B1:B10 中第一个条件格式应用到的单元格区域的地址。*/
function test() {
    let formatCondition = ActiveSheet.Range("B1:B10").FormatConditions.Item(1)
    console.log(formatCondition.AppliesTo.Address())
}
```
