---
title: "SetFirstPriority 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/%E6%96%B9%E6%B3%95/SetFirstPriority%20%E6%96%B9%E6%B3%95.html
---
# SetFirstPriority (方法)

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。

## 说明

当工作表中有多个条件格式规则时，如果之前未将此规则设置为优先级“1”，此方法将导致工作表上的所有其他现有规则的优先级都增加 1。

将在工作表级别应用条件格式规则的优先级。

## 示例

```javascript
/*本示例将活动工作表上区域 C1:C10 中第二个色阶条件格式的优先级值设置为“1”，并删除该区域上第一个色阶条件格式。*/
function test() {
    let colorScale = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    colorScale.SetFirstPriority()
    ActiveSheet.Range("C1:C10").FormatConditions.Item(1).Delete()
}
```

```javascript
/*本示例将工作表 Sheet1 上区域 C1:C10 中第二个色阶条件格式的优先级设置为“1”。*/
function test() {
    let colorScale = Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(2)
    colorScale.SetFirstPriority()
}
```
