---
title: "Forecast 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Forecast%20%E6%96%B9%E6%B3%95.html
---
# Forecast (方法)

根据已有的数值计算或预测未来值。此预测值为基于给定的 x 值推导出的 y 值。已知的数值为已有的 x 值和 y 值，再利用线性回归对新值进行预测。可以使用该函数对未来销售额、库存需求或消费趋势进行预测。

## 说明

- 如果 x 为非数值型，则 FORECAST 将返回错误值 #VALUE!。
- 如果 known_y's 和 known_x's 为空或含有不同个数的数据点，则 FORECAST 将返回错误值 #N/A。
- 如果 known_x's 的方差等于零，则 FORECAST 将返回错误值 #DIV/0!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | X - 要预测其值的数据点。 |
| Arg2 | any | 必填 | Known_y's - 相关数据数组或数据区域。 |
| Arg3 | any | 必填 | Known_x's - 独立数据数组或数据区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Forecast 方法根据现有的数组arr1，arr2和arr3计算或预测未来值，并分别赋值给A1和A2。*/
function test() {
    let arr1 = [0, 10, 20, 30, 40]
    let arr2 = [-9, -5, 40, 0, 100]
    let arr3 = [-13, 15, -43, 15, 100]
    Range("A1").Value2 = WorksheetFunction.Forecast(2, arr1, arr2)
    Range("A2").Value2 = WorksheetFunction.Forecast(5, arr3, arr2)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 Forecast 方法计算或预测未来值。*/
function test() {
    Range("C1").Value2 = -312
    Range("C2").Value2 = -236
    Range("D1").Value2 = 626
    Range("D2").Value2 = 928
    console.log(WorksheetFunction.Forecast(-3, Range("C1:C2"), Range("D1:D2")))
}
```
