---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConditionValue/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: bsyrrm1W
---

# Type (属性)

返回 **[XlConditionValueTypes](./../../枚举/XlConditionValueTypes 枚举.html)** 枚举的常量之一，该常量指定如何确定数据栏、色标或图标集条件格式的阈值。只读。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个（数据条）条件格式的最短的数据条的类型是否为xlConditionValueAutomaticMin。*/
function test() {
    let conditionValue = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).MinPoint
    console.log(conditionValue.Type == xlConditionValueAutomaticMin)
}
```

```javascript
/*本示例判断第一张工作表上区域 A1:A10 中第二个（数据条）条件格式的最长的数据条的类型，并通知用户。*/
function test() {
    let conditionValue = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(2).MaxPoint
    if (conditionValue.Type == xlConditionValueHighestValue) {
        console.log("最长数据条的类型为xlConditionValueHighestValue")
    } else {
        console.log("最长数据条的类型不为xlConditionValueHighestValue")
    }
}
```
