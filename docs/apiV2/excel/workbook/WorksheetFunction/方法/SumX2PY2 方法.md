---
title: "SumX2PY2 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/SumX2PY2%20%E6%96%B9%E6%B3%95.html
---
# SumX2PY2 (方法)

返回两数组中对应数值的平方和之和，平方和之和在统计计算中经常使用。

## 说明

- 参数应为数字，或者是包含数字的名称、数组或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果 array_x 和 array_y 中值的数目不同，则 SUMX2PY2 将返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Array_x - 第一个数值数组或区域。 |
| Arg2 | any | 必填 | Array_y - 第二个数值数组或区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 SumX2PY2 方法计算两个数组中相应值的平方和之和，并赋值给C1单元格。*/
function test() {
    let arr1 = [10, 20, 30.2]
    let arr2 = [10.3, -3.2, 6.2]
    Range("C1").Value2 = WorksheetFunction.SumX2PY2(arr1, arr2)
}
```

```javascript
/*本示例为D1等单元格分别赋值，使用 SumX2PY2 方法计算两个数组中相应值的平方和之和。*/
function test() {
    Range("D1").Value2 = 23.3
    Range("D2").Value2 = 20.03
    Range("D3").Value2 = 20.6
    Range("D4").Value2 = 21
    Range("E1").Value2 = -5
    Range("E2").Value2 = 2
    Range("E3").Value2 = 6
    Range("E4").Value2 = 3
    console.log(WorksheetFunction.SumX2PY2(Range("D1:D4"), Range("E1:E4")))
}
```
