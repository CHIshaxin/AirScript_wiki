---
title: "DateOperator 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/DateOperator%20%E5%B1%9E%E6%80%A7.html
---
# DateOperator (属性)

指定格式条件中所用的日期运算符。可读/写。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上区域 B1:B10 中第一个条件格式所用的日期运算符是否为xlThisWeek。*/
function test() {
    let formatCondition = Application.Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions.Item(1)
    console.log(formatCondition.DateOperator == xlThisWeek)
}
```

```javascript
/*本示例设置活动工作表上区域 B1:B10 中第一个条件格式所用的日期运算符。*/
function test() {
    let formatCondition = ActiveSheet.Range("B1:B10").FormatConditions.Item(1)
    formatCondition.DateOperator = xlToday
}
```
