---
title: "AveDev 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/AveDev%20%E6%96%B9%E6%B3%95.html
---
# AveDev (方法)

返回多个数据点与其平均值的绝对偏差的平均值。AveDev 是对数据集中可变性的度量。

## 说明

- AveDev 受输入数据中的度量单位的影响。
- 参数要么是数字，要么是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 平均偏差的计算公式为：

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - 要计算其绝对偏差平均值的 1 到 30 个参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用AveDev方法计算多个数据点与其平均值的绝对偏差的平均值，并将结果赋值于A5等单元格。*/
function test() {
    Range("A5").Value2 = Application.WorksheetFunction.AveDev(14, 37, 54, -20, 13)
    Range("B5").Value2 = Application.WorksheetFunction.AveDev(1, 3.5, 7.8, 9)
    Range("C5").Value2 = Application.WorksheetFunction.AveDev(2, 4, 6, 8.88, 10.1, 12.2, 14, 16)
}
```

```javascript
/*本示例为E3等单元格分别赋值，并使用AveDev方法计算多个数据点与其平均值的绝对偏差的平均值。*/
function test() {
    Range("E3").Value2 = 5
    Range("E4").Value2 = 25
    Range("E5").Value2 = -14.9
    Range("E6").Value2 = 43
    let avedev1 = Application.WorksheetFunction.AveDev(Range("E3:E6").Value2)
    console.log(avedev1)
}
```
