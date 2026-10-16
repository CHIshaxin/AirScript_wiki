---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLineCells/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Bk2XHqct
---

# Item (方法)

按 **[PivotLineCells](./../../PivotLineCells/PivotLineCells 对象.html)** 集合对象在集合中的位置返回特定元素。 此为只读属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 指定集合的元素的位置。 |

## 示例

```javascript
/*本示例显示活动工作表上数据透视表行轴上第一条数据透视线的PivotLineCells属性的第一个PivotCell的地址。*/
function test() {
    let pvtCell = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines(1).PivotLineCells.Item(1)
    console.log(pvtCell.Range.Address())
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表列轴上第一条数据透视线的PivotLineCells属性的第二个PivotCell对应的数据透视表字段的值。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotColumnAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Item(2).DataField.Value)
}
```
