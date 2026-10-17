---
title: "SetFirstPriority (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UniqueValues/%E6%96%B9%E6%B3%95/SetFirstPriority%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BdwI4w6M
---

# SetFirstPriority (方法)

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。

## 说明

当工作表中有多个条件格式规则时，如果之前未将此规则设置为优先级“1”，此方法将导致工作表上的所有其他现有规则的优先级都增加 1。

将在工作表级别应用条件格式规则的优先级。

## 示例

```javascript
/*本示例将活动工作表上区域 B1:B10 中第二个（UniqueValues）条件格式的优先级值设置为“1”。*/
function test() {
    let uniqueValues = ActiveSheet.Range("B1:B10").FormatConditions.Item(2)
    uniqueValues.SetFirstPriority()
}
```

```javascript
/*本示例将工作表 Sheet1 上区域 B1:B10 中第二个（UniqueValues）条件格式的优先级设置为“1”，然后设置该区域上第一个（UniqueValues）条件格式所应用于的数字格式。*/
function test() {
    let formatConditions = Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions
    formatConditions.Item(2).SetFirstPriority()
    formatConditions.Item(1).NumberFormat = "0.00"
}
```
