---
title: "FormatConditions (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/FormatConditions%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CMBCALyV
---

# FormatConditions (对象)

代表一个区域内所有条件格式的集合。

## 说明

**[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合可以包含多个条件格式。每个格式由一个 **[FormatCondition](./../FormatCondition/FormatCondition 对象.html)** 对象代表。

有关条件格式的详细信息，请参阅 **[FormatCondition](./../FormatCondition/FormatCondition 对象.html)** 对象。

使用 **FormatConditions** 属性可返回 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 对象。使用 **Add** 方法可新建条件格式，使用 **Modify** 方法可更改现有的条件格式。

## 示例

```javascript
/*本示例向第一张工作表上单元格区域 E1:E10 中添加条件格式*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Add(xlCellValue, xlGreater, "=$A$1")
    let boders = formatCondition.Borders
    boders.LineStyle = xlContinuous
    boders.Weight = xlThin
    boders.ColorIndex = 6
    let font = formatCondition.Font
    font.Bold = true
    font.ColorIndex = 3
}
```

```javascript
/*本示例在活动工作表上单元格区域 F1:F10 新增Top10条件格式。*/
function test() {
    let top = ActiveSheet.Range("F1:F10").FormatConditions.AddTop10()
    top.TopBottom = xlTop10Top
    top.Rank = 5
    top.Percent = false
    top.Font.Bold = true
    top.Interior.ColorIndex = 5
}
```
