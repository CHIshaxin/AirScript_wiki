---
title: "PivotColumnLine (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/PivotColumnLine%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CQ3Xuwpd
---

# PivotColumnLine (属性)

在列上返回特定 **[PivotCell](./../../PivotCell/PivotCell 对象.html)** 对象的 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 。只读 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 类型。

## 说明

如果 PivotCell 在行上，则 **PivotColumnLine** 属性返回一个运行时错误。

如果 PivotCell 在列上，则 **PivotColumnLine** 属性返回列 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 对象。

如果 PivotCell 在数据区域中，则 **PivotColumnLine** 属性返回对应的列 **[PivotLine](./../../PivotLine/PivotLine 对象.html)** 对象。

## 示例

```javascript
/*本示例显示包含活动单元格的数据透视项列轴上数据透视线的类型是否为“xlPivotLineRegular”。*/
function test() {
    console.log(ActiveCell.PivotCell.PivotColumnLine.LineType == xlPivotLineRegular)
}
```

```javascript
/*本示例显示包含单元格 M3 的数据透视项列轴上的位置。*/
function test() {
    let pvtLine = Range("M3").PivotCell.PivotColumnLine
    console.log(pvtLine.Position)
}
```
