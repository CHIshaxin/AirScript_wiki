---
title: "Kurt (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Kurt%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DQ70GbXP
---

# Kurt (方法)

返回一组数据的峰值。峰值反映与正态分布相比某一分布的尖锐度或平坦度。正峰值表示相对尖锐的分布。负峰值表示相对平坦的分布。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的数字的逻辑值和文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果数据点少于四个，或样本标准偏差等于零，则 KURT 返回错误值 #DIV/0!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1, number2, ... - 要计算其峰值的 1 到 30 个参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Kurt 方法计算一组数据的峰值，并分别赋值给D3和D4单元格。*/
function test() {
    Range("D3").Value2 = WorksheetFunction.Kurt(1, 2, 3, 4, 5)
    Range("D4").Value2 = WorksheetFunction.Kurt([1.1, 2.3, 6.5, 4, 8.6])
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 Kurt 方法计算区域C1:C4的峰值。*/
function test() {
    Range("C1").Value2 = 10.35
    Range("C2").Value2 = 25.3
    Range("C3").Value2 = 4.32
    Range("C4").Value2 = 5
    console.log(WorksheetFunction.Kurt(Range("C1:C4")))
}
```
