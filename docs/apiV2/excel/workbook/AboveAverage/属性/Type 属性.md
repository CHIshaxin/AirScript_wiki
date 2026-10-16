---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

返回 **[XlFormatConditionType](./../../枚举/XlFormatConditionType 枚举.html)** 枚举的常量之一，该常量指定条件格式的类型。只读。

## 说明

此属性将始终返回 **Long** 值“12”，等价于 **xlAboveAverageCondition**。

## 示例

```javascript
/*本示例显示活动工作表上区域 C1:C10 中第二个条件格式的类型是否为xlAboveAverageCondition。*/
function test() {
    let aboveAverage = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    console.log(aboveAverage.Type == xlAboveAverageCondition)
}
```

```javascript
/*本示例判断活动工作表上区域 C1:C10 中第一个条件格式的类型是否为xlAboveAverageCondition，并通知用户。*/
function test() {
    let formatCondition = ActiveSheet.Range("C1:C10").FormatConditions.Item(1)
    if (formatCondition.Type == xlAboveAverageCondition) {
        console.log("类型为xlAboveAverageCondition")
    } else {
        console.log("类型不为xlAboveAverageCondition")
    }
}
```
