---
title: "Ddb 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Ddb%20%E6%96%B9%E6%B3%95.html
---
# Ddb (方法)

使用双倍余额递减法或其他指定方法返回某项资产在指定周期内的折旧值。

## 说明

这五个参数都必须是正数。

- 双倍余额递减法以加速的速率计算折旧值。折旧值在第一个周期是最高的，在后续周期中会减少。DDB 使用下面的公式计算一个周期的折旧值： Min( (cost - total depreciation from prior periods) \* (factor/life), (cost - salvage - total depreciation from prior periods) )
- 如果不想使用双倍余额递减法，则请改变 factor 的值。
- 当折旧值大于余额递减计算值时，如果想转换到直线折旧法，则请使用 VDB 函数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Cost - 资产的初始成本。 |
| Arg2 | double | 必填 | Salvage - 资产在折旧期末的价值（有时也称为资产残值）。此值可为 0。 |
| Arg3 | double | 必填 | Life - 资产进行折旧计算的周期数（有时也称作资产的生命周期）。 |
| Arg4 | double | 必填 | Period - 要为其计算折旧值的周期。Period 必须与 life 使用相同的单位。 |
| Arg5 | any | 可选 | Factor - 余额递减速率。如果省略 factor，则假定其值为 2（双倍余额递减法）。 |

## 返回值

Double

## 示例

```javascript
/*本示例为C3等单元格分别赋值，并使用Ddb方法计算指定期间内资产的折旧值。*/
function test() {
    Range("C3").Value2 = 250000
    Range("C4").Value2 = 3000
    Range("C5").Value2 = 4
    let ddb1 = Application.WorksheetFunction.Ddb(Range("C3").Value2, Range("C4").Value2, Range("C5").Value2, 1, 4)
    console.log(ddb1)
}
```

```javascript
/*本示例使用Ddb方法计算指定期间内资产的折旧值，并将结果赋值于B3等单元格。*/
function test() {
    Range("B3").Value2 = Application.WorksheetFunction.Ddb(350000, 5000, 8, 5)
    Range("B4").Value2 = Application.WorksheetFunction.Ddb(4000000, 70000, 3, 1, 1.5)
    Range("B5").Value2 = Application.WorksheetFunction.Ddb(80000, 4000, 5, 3, 3)
}
```
