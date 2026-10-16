---
title: "StopIfTrue (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/StopIfTrue%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B1nSxBH9
---

# StopIfTrue (属性)

返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。

## 说明

为了支持向后兼容性，此属性的默认值为 **True**，而这与在用户界面中创建的格式规则的行为相反。

## 示例

```javascript
/*本示例判断工作表 Sheet1 上区域 A1:A10 中第一个条件格式的计算结果为True时是否不计算单元格上的其他格式规则，并通知用户。*/
function test() {
    let formatCondition = Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(1)
    if (formatCondition.StopIfTrue) {
        console.log("条件格式的计算结果为 True 时不计算单元格上的其他格式规则")
    } else {
        console.log("条件格式的计算结果为 True 时计算单元格上的其他格式规则")
    }
}
```

```javascript
/*本示例将活动工作表上区域 A1:A10 中第一个条件格式设置为当计算结果为True时不计算单元格上的其他格式规则。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.StopIfTrue = true
}
```
