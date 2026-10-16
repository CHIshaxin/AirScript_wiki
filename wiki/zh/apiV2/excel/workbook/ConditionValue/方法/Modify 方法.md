---
title: "Modify (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConditionValue/%E6%96%B9%E6%B3%95/Modify%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BfPSKYEN
---

# Modify (方法)

修改数据条条件格式规则计算最长数据条或最短数据条的方法。

## 说明

下表描述了每种计算类型的可接受阈值。

| NewType 参数 | NewValue 参数 |
| --- | --- |
| xlConditionLowestValue | 忽略参数 |
| xlConditionHighestValue | 忽略参数 |
| xlConditionValueNumber | 任何数字 |
| xlConditionValuePercent | 0 到 100 之间的任何数字 |
| xlConditionValuePercentile | 0 到 100 之间的任何数字 |
| xlConditionValueFormula | 一个返回单一数字的公式 |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |

## 示例

```javascript
/*本示例修改活动工作表上区域 A1:A10 中第一个（数据条）条件格式的最长数据条的计算方法。*/
function test() {
    let conditionValue = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).MaxPoint
    conditionValue.Modify(xlConditionValuePercent, 200)
}
```

```javascript
/*本示例修改第一张工作表上区域 A1:A10 中第二个（数据条）条件格式的最短的数据条的计算方法。*/
function test() {
    let conditionValue = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(2).MinPoint
    conditionValue.Modify(xlConditionValueNumber, -1)
}
```
