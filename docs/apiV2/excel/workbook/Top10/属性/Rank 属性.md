---
title: "Rank 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Top10/%E5%B1%9E%E6%80%A7/Rank%20%E5%B1%9E%E6%80%A7.html
---
# Rank (属性)

返回或设置一个 **Long** 值，为条件格式规则指定排位值的数值或百分比。

## 示例

```javascript
/*本示例显示活动工作表上区域 C1:C10 中第二个（Top10）条件格式的排位值的数值。*/
function test() {
    let top = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    console.log(top.Rank)
}
```

```javascript
/*本示例将第一张工作表上区域 C1:C10 中第一个（Top10）条件格式设置为按百分比值确定排位，并设置该条件格式的排位值的百分比。*/
function test() {
    let top = Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(1)
    top.Percent = true
    top.Rank = 30
}
```
