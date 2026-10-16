---
title: "LineType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLine/%E5%B1%9E%E6%80%A7/LineType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BCcbUy3D
---

# LineType (属性)

返回一个表示 PivotLine 类型的 **[XlPivotLineType](./../../枚举/XlPivotLineType 枚举.html)** 常量。只读。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第一条数据透视线的类型是否为“xlPivotLineRegular”。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(1)
    console.log(pvtLine.LineType == xlPivotLineRegular)
}
```

```javascript
/*本示例判断工作表 Sheet1 中第一张数据透视表中行轴上第二条数据透视线的类型是否为“xlPivotLineGrandTotal”，并通知用户。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(2)
    if (pvtLine.LineType == xlPivotLineGrandTotal) {
        console.log("数据透视线的类型是 xlPivotLineGrandTotal")
    } else {
        console.log("数据透视线的类型不是 xlPivotLineGrandTotal")
    }
}
```
