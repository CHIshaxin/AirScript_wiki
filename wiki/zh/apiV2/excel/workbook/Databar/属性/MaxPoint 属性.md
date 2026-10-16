---
title: "MaxPoint (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/MaxPoint%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: mj8nSpgU
---

# MaxPoint (属性)

返回一个 **[ConditionValue](./../../ConditionValue/ConditionValue 对象.html)** 对象，该对象指定如何针对数据条条件格式对最长的数据条进行求值。

## 说明

通过使用 **Type** 和 **Value** 属性，您可以指定求值的类型和值。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个为数据条的条件格式的最长数据条的值。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.MaxPoint.Value)
}
```

```javascript
/*本示例显示活动单元格上第一个为数据条的条件格式的最长数据条的类型是否为xlConditionValueNumber。*/
function test() {
    let databar = ActiveCell.FormatConditions.Item(1)
    console.log(databar.MaxPoint.Type == xlConditionValueNumber)
}
```
