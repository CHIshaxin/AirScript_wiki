---
title: "Priority 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/%E5%B1%9E%E6%80%A7/Priority%20%E5%B1%9E%E6%80%A7.html
---
# Priority (属性)

返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。

## 说明

设置优先级时，值必须为介于 1 和工作表上条件格式规则总数之间的正整数。对于工作表上的所有规则，优先级必须为唯一值，因此，如果更改指定条件格式规则的优先级，将可能会导致工作表上其他规则的优先级值移位。

## 示例

```javascript
/*本示例显示活动工作表上区域 C1:C10 中第二个色阶条件格式的优先级值。*/
function test() {
    let colorScale = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    console.log(colorScale.Priority)
}
```

```javascript
/*本示例设置工作表 Sheet1 上区域 C1:C10 中第一个色阶条件格式的优先级值。*/
function test() {
    let colorScale = Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    colorScale.Priority = 2
}
```
