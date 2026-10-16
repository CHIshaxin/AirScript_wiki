---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CXm9KZEI
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*此示例为第一张工作表上 E1:E10 单元格区域的现有条件格式设置格式属性。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    let boders = formatCondition.Borders
    boders.LineStyle = xlContinuous
    boders.Weight = xlThin
    boders.ColorIndex = 6
}
```

```javascript
/*本示例选中工作表 Sheet1 上区域 B1:B10 中第二个条件格式应用到的单元格区域。*/
function test() {
    let formatCondition = Application.Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions.Item(2)
    formatCondition.AppliesTo.Select()
}
```
