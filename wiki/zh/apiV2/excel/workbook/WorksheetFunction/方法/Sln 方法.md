---
title: "Sln (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Sln%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BY465_Dx
---

# Sln (方法)

返回某项资产一个周期的直线折旧值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Cost - 资产的初始成本。 |
| Arg2 | double | 必填 | Salvage - 资产在折旧期末的价值（有时也称为资产残值）。 |
| Arg3 | double | 必填 | Life - 资产进行折旧计算的周期数（有时也称作资产的生命周期）。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Sln 方法计算资产3年期和5年期的直线折旧值，并分别赋值给A1和A2单元格。*/
function test() {
    Range("C3").Value2 = WorksheetFunction.Sln(100, 10000, 3)
    Range("C4").Value2 = WorksheetFunction.Sln(10000, 100, 5)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Sln 方法计算资产4年期的直线折旧值。*/
function test() {
    Range("B1").Value2 = 9234
    Range("C1").Value2 = 6548
    Range("D1").Value2 = 4
    console.log(WorksheetFunction.Sln(Range("B1").Value2, Range("C1").Value2, Range("D1").Value2))
}
```
