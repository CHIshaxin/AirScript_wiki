---
title: "FormatCondition 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/FormatCondition%20%E5%AF%B9%E8%B1%A1.html
---
# FormatCondition (对象)

代表条件格式。

## 说明

**[FormatCondition](./../FormatCondition/FormatCondition 对象.html)** 对象是 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合的成员。对于给定区域， **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合中包含的条件格式不能超过三个。

使用 **Add** 方法可新建条件格式。如果区域内存在多种格式，则可使用 **Modify** 方法更改其中一种格式，或使用 **Delete** 方法删除一种格式，然后使用 **Add** 方法创建一种新格式。

使用 **[FormatCondition](./../FormatCondition/FormatCondition 对象.html)** 对象的 **Font**、**Borders** 和 **Interior** 属性可控制已设置格式的单元格的外观。条件格式对象模型不支持这些对象的某些属性。下表列出所有可与条件格式一起使用的属性。

| 对象 | 属性 |
| --- | --- |
| **Font** | **Bold** **Color** **ColorIndex** **FontStyle** **Italic** **Strikethrough** **Underline** 无法使用会计用下划线样式。 |
| **Border** | **Bottom** **Color** **Left** **Right** **Style** 可使用下列边框样式（其他均不可用）：**xlNone**、**xlSolid**、**xlDash**、**xlDot**、**xlDashDot**、**xlDashDotDot**、**xlGray50**、**xlGray75** 和 **xlGray25**。 **Top** **Weight** 可使用下列边框粗细（其他均不可用）：**xlWeightHairline** 和 **xlWeightThin**。 |
| **Interior** | **Color** **ColorIndex** **Pattern** **PatternColorIndex** |

使用 **FormatConditions**(*index*)（其中 *index* 为条件格式的索引号）可返回 **[FormatCondition](./../FormatCondition/FormatCondition 对象.html)** 对象。

## 示例

```javascript
/*本示例设置第一张工作表上 E1:E10 单元格的现有条件格式的格式属性。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
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
/*本示例设置活动工作表上单元格区域 A1:A10 的第一个条件格式所应用于的单元格区域，并修改现有条件格式。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.ModifyAppliesToRange(Range("A2:A9"))
    formatCondition.Modify(xlCellValue, xlBetween, "=4", "=7")
}
```
