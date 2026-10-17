---
title: "Log10 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Log10%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CoIlx373
---

# Log10 (方法)

返回以 10 为底的对数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要计算其以 10 为底的对数的正实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Log10 方法分别计算数值50和0.6以10为底的对数，并分别赋值给A4和A5单元格。*/
function test() {
    Range("A4").Value2 = WorksheetFunction.Log10(50)
    Range("A5").Value2 = WorksheetFunction.Log10(0.6)
}
```

```javascript
/*本示例为B1单元格赋值，使用 Log10 方法计算B1单元格的值以10为底的对数。*/
function test() {
    Range("B1").Value2 = 123
    console.log(WorksheetFunction.Log10(Range("B1").Value2))
}
```
