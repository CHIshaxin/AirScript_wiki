---
title: "MinPoint 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/MinPoint%20%E5%B1%9E%E6%80%A7.html
---
# MinPoint (属性)

返回一个 **[ConditionValue](./../../ConditionValue/ConditionValue 对象.html)** 对象，该对象指定如何针对数据条条件格式对最短的数据条进行求值。

## 说明

通过使用 **Type** 和 **Value** 属性，您可以指定求值的类型和值。

## 示例

```javascript
/*本示例判断如果活动单元格上第一个条件格式（数据条）的最短数据条的类型为xlConditionValuePercent，则显示该最短数据条的值。*/
function test() {
    let databar = ActiveCell.FormatConditions.Item(1)
    if (databar.MinPoint.Type == xlConditionValuePercent) {
        console.log(databar.MinPoint.Value)
    }
}
```

```javascript
/*本示例修改活动工作表上区域 A1:A10 中第一个条件格式（数据条）的最短的数据条的规则。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.MinPoint.Modify(xlConditionValueNumber, -1)
}
```
