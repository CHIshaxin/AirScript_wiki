---
title: "AppliesTo (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E5%B1%9E%E6%80%A7/AppliesTo%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BfH7Az-Q
---

# AppliesTo (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，指定格式规则将应用于的单元格区域。

## 示例

```javascript
/*本示例显示第一张工作表上区域 A1:A10 中第一个图标集条件格式应用到的单元格区域。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    console.log(iconSet.AppliesTo.Address())
}
```

```javascript
/*本示例选中工作表 Sheet1 上区域 A1:A10 中第二个图标集条件格式应用到的单元格区域。*/
function test() {
    let iconSet = Application.Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(2)
    iconSet.AppliesTo.Select()
}
```
