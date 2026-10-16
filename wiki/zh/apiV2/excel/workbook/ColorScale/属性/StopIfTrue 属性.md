---
title: "StopIfTrue (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/%E5%B1%9E%E6%80%A7/StopIfTrue%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BrHcKq2A
---

# StopIfTrue (属性)

返回一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则，只读。

## 说明

为了支持向后兼容性，此属性的默认值为 **True**，而这与在用户界面中创建的格式规则的行为相反。

## 示例

```javascript
/*本示例判断工作表 Sheet1 上区域 C1:C10 中第一个色阶条件格式的计算结果为True时是否停止计算单元格上的其他格式规则，并通知用户。*/
function test() {
    let colorScale = Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    if (colorScale.StopIfTrue) {
        console.log("条件格式的计算结果为 True 时停止计算单元格上的其他格式规则")
    } else {
        console.log("条件格式的计算结果为 True 时计算单元格上的其他格式规则")
    }
}
```

```javascript
/*本示例显示活动工作表上区域 C1:C10 中第二个色阶条件格式的计算结果为True时是否停止计算单元格上的其他格式规则。*/
function test() {
    let colorScale = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    console.log(colorScale.StopIfTrue)
}
```
