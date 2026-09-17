---
title: "GeStep 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/GeStep%20%E6%96%B9%E6%B3%95.html
---
# GeStep (方法)

如果 Number 大于等于 step，返回 1，否则返回 0（零）。使用此函数可对一组值进行筛选。例如，通过计算多个函数 GESTEP 的返回值，可以检测出数据集中超过某个临界值的数据个数。

## 说明

如果任一参数为非数值型，则 GESTEP 将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 要依据 step 检验的值。 |
| Arg2 | any | 可选 | Step - 阈值。如果省略 step 的值，则 GESTEP 将使用零。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 GeStep 方法判断 Number 参数是否大于等于 Step 参数，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.GeStep(13, 15)
    Range("B2").Value2 = WorksheetFunction.GeStep(10)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 GeStep 方法判断C1单元格的值是否大于等于D1单元格的值。*/
function test() {
    Range("C1").Value2 = 45
    Range("D1").Value2 = 36
    console.log(WorksheetFunction.GeStep(Range("C1").Value2, Range("D1").Value2))
}
```
