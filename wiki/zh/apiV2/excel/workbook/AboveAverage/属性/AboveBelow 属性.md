---
title: "AboveBelow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/%E5%B1%9E%E6%80%A7/AboveBelow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CExMNR7W
---

# AboveBelow (属性)

返回或设置 **[XlAboveBelow](./../../枚举/XlAboveBelow 枚举.html)** 枚举的常量之一，指定条件格式规则是否查找高于或低于区域的平均值或标准偏差的单元格值。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上区域 B1:B10 中第一个（AboveAverage）条件格式的AboveBelow属性是否为xlAboveAverage。*/
function test() {
    let aboveAverage = Application.Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions.Item(1)
    console.log(aboveAverage.AboveBelow == xlAboveAverage)
}
```

```javascript
/*本示例设置活动工作表上区域 B1:B10 中第一个（AboveAverage）条件格式的AboveBelow属性。*/
function test() {
    let aboveAverage = ActiveSheet.Range("B1:B10").FormatConditions.Item(1)
    aboveAverage.AboveBelow = xlEqualBelowAverage
}
```
