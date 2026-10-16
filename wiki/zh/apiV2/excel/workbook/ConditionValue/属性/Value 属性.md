---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConditionValue/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CsSG4n2V
---

# Value (属性)

针对数据条条件格式返回最短或最长的数据条的阈值。只读 **Variant** 类型。

## 示例

```javascript
/*本示例显示第一张工作表上区域 A1:A10 中第二个（数据条）条件格式的最长的数据条的阈值。*/
function test() {
    let conditionValue = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(2).MaxPoint
    console.log(conditionValue.Value)
}
```

```javascript
/*本示例判断如果活动工作表上区域 A1:A10 中第一个（数据条）条件格式的最短的数据条的类型为xlConditionValueNumber，则显示该数据条的阈值。*/
function test() {
    let conditionValue = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).MinPoint
    if (conditionValue.Type == xlConditionValueNumber) {
        console.log(conditionValue.Value)
    }
}
```
