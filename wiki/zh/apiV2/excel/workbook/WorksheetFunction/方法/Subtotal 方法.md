---
title: "Subtotal (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Subtotal%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: XVa1mH0k
---

# Subtotal (方法)

创建分类汇总。返回 Double 表示小计。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 一个数字，表示要由该方法使用的聚合函数。 |
| Arg2 | Range | 必填 | 要为其计算分类汇总的第一个 **[Range](./../../Range/Range 对象.html)** 对象。 |

## 返回值

Double

## 示例

```javascript
/*本示例为C2等单元格分别赋值，使用 Subtotal 方法创建分类汇总，并分别赋值给B1和B2单元格。*/
function test() {
    Range("C2").Value2 = 0.3
    Range("C3").Value2 = 10
    Range("C4").Value2 = 11
    Range("C5").Value2 = 23.3
    Range("B1").Value2 = WorksheetFunction.Subtotal(1, Range("C2:C5"))
    Range("B2").Value2 = WorksheetFunction.Subtotal(3, Range("C2:C5"))
}
```

```javascript
/*本示例为A2等单元格分别赋值，使用 Subtotal 方法创建分类汇总。*/
function test() {
    Range("A2").Value2 = 0.3
    Range("A3").Value2 = 4.3
    Range("A4").Value2 = 1.3
    Range("A5").Value2 = 2.3
    console.log(WorksheetFunction.Subtotal(9, Range("A2:A5")))
}
```
