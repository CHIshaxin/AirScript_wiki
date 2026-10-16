---
title: "SetFirstPriority (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/%E6%96%B9%E6%B3%95/SetFirstPriority%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C2LfDVHD
---

# SetFirstPriority (方法)

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。

## 说明

当工作表中有多个条件格式规则时，如果之前未将此规则设置为优先级“1”，此方法将导致工作表上的所有其他现有规则的优先级都增加 1。

将在工作表级别应用条件格式规则的优先级。

## 示例

```javascript
/*本示例将活动工作表上区域 C1:C10 中第二个（AboveAverage）条件格式的优先级值设置为“1”。*/
function test() {
    let aboveAverage = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    aboveAverage.SetFirstPriority()
}
```

```javascript
/*本示例将工作表 Sheet1 上区域 C1:C10 中第一个（AboveAverage）条件格式的优先级设置为“1”，并显示区域 C7:C10 中第一个条件格式的优先级值。*/
function test() {
    let aboveAverage = Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    aboveAverage.SetFirstPriority()
    console.log(Range("C7:C10").FormatConditions.Item(1).Priority)
}
```
