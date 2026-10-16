---
title: "ConsolidationFunction 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/ConsolidationFunction%20%E5%B1%9E%E6%80%A7.html
---
# ConsolidationFunction (属性)

返回当前合并计算所使用的函数代码。可以是 **[XlConsolidationFunction](./../../枚举/XlConsolidationFunction 枚举.html)** 常量之一。**Long** 类型，只读。

## 示例

```javascript
/*如果当前合并计算所使用的函数为 SUM 函数，本示例将显示一消息框。*/
function test() {
    if (Worksheets.Item("Sheet1").ConsolidationFunction == xlSum) {
        console.log("Sheet1 使用 SUM 函数合并。")
    }
}
```

```javascript
/*本示例判断如果当前合并计算所使用的函数不是 Count 函数，则通知用户活动工作表合并计算所使用的函数代码。*/
function test() {
    if (ActiveSheet.ConsolidationFunction != xlCount) {
        console.log(ActiveSheet.ConsolidationFunction)
    }
}
```
