---
title: "StopIfTrue 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/StopIfTrue%20%E5%B1%9E%E6%80%A7.html
---
# StopIfTrue (属性)

返回一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则，只读。

## 说明

为了支持向后兼容性，此属性的默认值是 **False**，而这与在用户界面中创建的格式规则的行为相反。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个条件格式（数据条）的计算结果为True时是否不计算单元格上的其他格式规则。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.StopIfTrue)
}
```

```javascript
/*本示例判断工作表 Sheet1 上区域 A1:A8 中第一个条件格式（数据条）的计算结果为True时是否不计算单元格上的其他格式规则，并通知用户。*/
function test() {
    let databar = Worksheets.Item("Sheet1").Range("A1:A8").FormatConditions.Item(1)
    if (databar.StopIfTrue) {
        console.log("数据条的计算结果为 True 时不计算单元格上的其他格式规则")
    } else {
        console.log("数据条的计算结果为 True 时计算单元格上的其他格式规则")
    }
}
```
