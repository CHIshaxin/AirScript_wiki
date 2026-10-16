---
title: "FormatConditions (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/FormatConditions%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BjzQUEYv
---

# FormatConditions (属性)

返回一个 **[FormatConditions](./../../FormatConditions/FormatConditions 对象.html)** 集合，该集合表示指定区域的所有条件格式。只读。

## 示例

```javascript
/*本示例更改单元格区域 E1:E10 的现有条件格式。*/
function test() {
    Range("E1:E10").FormatConditions.Item(1).Modify(xlCellValue, xlLess, "=$A$1")
}
```

```javascript
/*本示例为 A1:C3 单元格区域添加条件格式，并显示 B2 单元格的条件格式数量，然后将 A1:C3 单元格区域条件格式删除，再显示 B2 单元格的条件格式数量*/
function test() {
    Range("A1:C3").FormatConditions.Add(xlUniqueValues, xlGreater, 6)
    console.log(Range("B2").FormatConditions.Count)
    Range("A1:C3").FormatConditions.Delete()
    console.log(Range("B2").FormatConditions.Count)
}
```
