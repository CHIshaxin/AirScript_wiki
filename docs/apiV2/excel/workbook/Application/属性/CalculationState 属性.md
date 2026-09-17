---
title: "CalculationState 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/CalculationState%20%E5%B1%9E%E6%80%A7.html
---
# CalculationState (属性)

返回一个 **[XlCalculationState](./../../枚举/XlCalculationState 枚举.html)** 常量，该常量针对 ET 中正在执行的任意计算指示应用程序的计算状态。只读。

## 示例

```javascript
/*在本示例中，ET 检查是否存在正在执行的计算。如果没有，则将计算状态消息显示为“Done”。否则就将其显示为“Not Done”。*/
function test() {
    if (Application.CalculationState == xlDone) {
        console.log("Done")
    } else {
        console.log("Not Done")
    }
}
```

```javascript
/*本示例检查ET是否在执行计算，并通知用户。*/
function test() {
    if (Application.CalculationState == xlCalculating) {
        console.log("正在计算！")
    }
}
```
