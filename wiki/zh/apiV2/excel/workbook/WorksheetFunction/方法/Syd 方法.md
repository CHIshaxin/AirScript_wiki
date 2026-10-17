---
title: "Syd (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Syd%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Q_bGflhU
---

# Syd (方法)

返回某项资产按年限总和折旧法计算的指定期间的折旧值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Cost - 资产的初始成本。 |
| Arg2 | double | 必填 | Salvage - 资产在折旧期末的价值（有时也称为资产残值）。 |
| Arg3 | double | 必填 | Life - 资产进行折旧计算的周期数（有时也称作资产的生命周期）。 |
| Arg4 | double | 必填 | Per - 期间，其单位必须与 life 相同。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Syd 方法计算某项资产按年限总和折旧法计算3年期和1年期的折旧值，并分别赋值给C1和C2单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.Syd(10000, 200, 3, 1)
    Range("C2").Value2 = WorksheetFunction.Syd(0, 100.23, 1, 1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Syd 方法计算某项资产按年限总和折旧法计算5年期的折旧值。*/
function test() {
    Range("A1").Value2 = 500.36
    Range("B1").Value2 = 200.03
    Range("C1").Value2 = 5
    Range("D1").Value2 = 1
    console.log(WorksheetFunction.Syd(Range("A1").Value2, Range("B1").Value2, Range("C1").Value2, Range("D1").Value2))
}
```
