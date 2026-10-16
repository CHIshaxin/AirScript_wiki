---
title: "SetLastPriority 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/%E6%96%B9%E6%B3%95/SetLastPriority%20%E6%96%B9%E6%B3%95.html
---
# SetLastPriority (方法)

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。

## 说明

优先级的实际值将等于工作表上条件格式规则的总数。当工作表中有多个条件格式规则时，此方法将导致优先级值大于此规则的规则的优先级降低 1。

将在工作表级别应用条件格式规则的优先级。

## 示例

```javascript
/*本示例设置活动工作表上区域 C1:C10 中第二个色阶条件格式的求值顺序。*/
function test() {
    let colorScale = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    colorScale.SetLastPriority()
}
```

```javascript
/*本示例设置工作表 Sheet1 上区域 C1:C10 中第一个色阶条件格式的求值顺序，并显示该条件格式的优先级值。*/
function test() {
    let colorScale = Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    colorScale.SetLastPriority()
    console.log(colorScale.Priority)
}
```
