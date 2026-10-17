---
title: "Db (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Db%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: lMHEcmLd
---

# Db (方法)

使用固定余额递减法返回指定周期内某项资产的折旧值。

## 说明

- 固定余额递减法用于计算固定利率下的资产折旧值，函数 DB 使用下列计算公式来计算一个期间的折旧值： (cost - 前期折旧总值 ) * rate 其中： rate = 1 - ((salvage / cost) ^ (1 / life))，保留三位小数
- 第一个周期和最后一个周期的折旧属于特例。对于第一个周期，函数 DB 的计算公式为： cost * rate * month / 12
- 对于最后一个周期，DB 的计算公式为： ((cost - 前期折旧总值) * rate * (12 - month)) / 12

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Cost - 资产的初始成本。 |
| Arg2 | double | 必填 | Salvage - 资产在折旧期末的价值（有时也称为资产残值）。 |
| Arg3 | double | 必填 | Life - 资产进行折旧计算的周期数（有时也称作资产的生命周期）。 |
| Arg4 | double | 必填 | Period - 要为其计算折旧值的周期。Period 必须与 life 使用相同的单位。 |
| Arg5 | any | 可选 | Month - 第一年的月份数。如果省略，则假定为 12。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用Db方法计算指定周期内某项资产的折旧值，并将结果赋值于B2等单元格。*/
function test() {
    Range("B2").Value2 = Application.WorksheetFunction.Db(200000, 3000, 5, 3, 7)
    Range("B3").Value2 = Application.WorksheetFunction.Db(1000000, 4500, 9, 7, 8)
    Range("B4").Value2 = Application.WorksheetFunction.Db(25000000, 30000, 8, 4, 2)
}
```

```javascript
/*本示例为A2等单元格分别赋值，并使用Db方法以A2单元格的值为初始成本、A3单元格的值为资产残值、A4单元格为周期来计算折旧值。*/
function test() {
    Range("A2").Value2 = 120000
    Range("A3").Value2 = 20000
    Range("A4").Value2 = 6
    let db1 = Application.WorksheetFunction.Db(Range("A2").Value2, Range("A3").Value2, Range("A4").Value2, 2, 1)
    console.log(db1)
}
```
