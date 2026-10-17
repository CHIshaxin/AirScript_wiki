---
title: "StEyx (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/StEyx%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: GR9jBpuW
---

# StEyx (方法)

返回通过线性回归法计算每个 x 的 y 预测值时所产生的标准误差。标准误差用来度量根据单个 x 变量计算出的 y 预测值的误差量。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，将会产生错误。
- 如果 known_y's 和 known_x's 的数据点个数不同，STEYX 将返回错误值 #N/A。
- 如果 known_y's 和 known_x's 为空或数据点个数少于三，STEYX 将返回错误值 #DIV/0!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Known_y's - 相关数据点的数组或区域。 |
| Arg2 | any | 必填 | Known_x's - 独立数据点的数组或区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 StEyx 方法获取通过线性回归法计算每个 x 的 y 预测值时所产生的标准误差，并分别赋值给C1和C2单元格。*/
function test() {
    let arr1 = [-1.1, 2.3, 4, 5.3, 6]
    let arr2 = [-0.9, 1.3, 3, 2.3, 4]
    Range("C1").Value2 = WorksheetFunction.StEyx(arr1, arr2)
    Range("C2").Value2 = WorksheetFunction.StEyx(arr2, arr1)
}
```

```javascript
/*本示例为A1等单元格赋值，使用 StEyx 方法获取通过线性回归法计算每个 x 的 y 预测值时所产生的标准误差。*/
function test() {
    Range("A1").Value2 = 1.2
    Range("A2").Value2 = 2.1
    Range("A3").Value2 = 3.2
    Range("A4").Value2 = 4.4
    Range("B1").Value2 = 3.4
    Range("B2").Value2 = 4.7
    Range("B3").Value2 = 5.4
    Range("B4").Value2 = 6.3
    console.log(WorksheetFunction.StEyx(Range("A1:A4"), Range("B1:B4")))
}
```
