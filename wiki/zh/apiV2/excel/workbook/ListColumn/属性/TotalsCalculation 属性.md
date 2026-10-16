---
title: "TotalsCalculation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumn/%E5%B1%9E%E6%80%A7/TotalsCalculation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: qxgVAzHB
---

# TotalsCalculation (属性)

基于 **[XlTotalsCalculation](./../../枚举/XlTotalsCalculation 枚举.html)** 枚举的值，确定列表列的“总计”行中的计算类型。可读写。

## 说明

| **[XlTotalsCalculation](./../../枚举/XlTotalsCalculation 枚举.html)** 可为以下 **[XlTotalsCalculation](./../../枚举/XlTotalsCalculation 枚举.html)** 常量之一。 |
| --- |
| **xlTotalsCalculationNone** |
| **xlTotalsCalculationSum** |
| **xlTotalsCalculationAverage** |
| **xlTotalsCalculationCount** |
| **xlTotalsCalculationCountNums** |
| **xlTotalsCalculationMin** |
| **xlTotalsCalculationStdDev** |
| **xlTotalsCalculationVar** |
| **xlTotalsCalculationMax** |

无须显示“总计”行即可设置此属性。此属性没有固定的“默认”值。随着对其他列的添加或删除，ET 会更改该属性的状态。

## 示例

```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第一列总计行的计算类型设置为xlTotalsCalculationSum。*/
function test() {
    Application.ActiveSheet.ListObjects.Item(1).ListColumns.Item(1).TotalsCalculation = xlTotalsCalculationSum
}
```

```javascript
/*本示例显示第一张工作表的第一个 ListObject 对象中第四列的总计行的计算类型是否为xlTotalsCalculationAverage。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.ListColumns.Item(4).TotalsCalculation == xlTotalsCalculationAverage)
}
```
