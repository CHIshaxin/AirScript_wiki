---
title: "PivotLines 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotAxis/%E5%B1%9E%E6%80%A7/PivotLines%20%E5%B1%9E%E6%80%A7.html
---
# PivotLines (属性)

返回附加到指定 **[PivotAxis](./../../PivotAxis/PivotAxis 对象.html)** 对象的 PivotLines。只读。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中行轴上数据透视线的数量。*/
function test() {
    let pvtLines = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines
    console.log(pvtLines.Count)
}
```

```javascript
/*本示例显示活动工作表上数据透视表中行轴上第二条数据透视线的类型是否为“xlPivotLineRegular”。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines.Item(2)
    console.log(pvtLine.LineType == xlPivotLineRegular)
}
```
