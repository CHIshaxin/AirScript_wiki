---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLines/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

根据 **[PivotLines](./../../PivotLines/PivotLines 对象.html)** 集合对象特定元素在集合中的位置返回该特定元素。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 指定集合的元素的位置。 |

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中列轴上第二条数据透视线的PivotLineCells属性第一个PivotCell对应数据透视表项的值。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines.Item(2)
    console.log(pvtLine.PivotLineCells.Item(1).PivotItem.Value)
}
```

```javascript
/*本示例显示活动工作表上数据透视表中行轴上第一条数据透视线的类型是否为“xlPivotLineRegular”。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines.Item(1)
    console.log(pvtLine.LineType == xlPivotLineRegular)
}
```
