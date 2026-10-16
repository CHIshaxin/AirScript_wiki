---
title: "TextOperator 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/TextOperator%20%E5%B1%9E%E6%80%A7.html
---
# TextOperator (属性)

返回或设置一个 **[XlContainsOperator](./../../枚举/XlContainsOperator 枚举.html)** 枚举常量，该常量指定条件格式规则执行的文本搜索。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个条件格式执行的文本搜索是否为xlContains。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.TextOperator == xlContains)
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式执行的文本搜索。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.TextOperator = xlDoesNotContain
}
```
