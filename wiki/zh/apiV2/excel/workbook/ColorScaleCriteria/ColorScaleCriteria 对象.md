---
title: "ColorScaleCriteria (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScaleCriteria/ColorScaleCriteria%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CbJ8Tpe4
---

# ColorScaleCriteria (对象)

代表色阶条件格式的所有条件的 **ColorScaleCriterion** 对象的集合。每个条件指定了色阶的最小、中点或最大阈值。

## 说明

要返回 **[ColorScaleCriteria](./../ColorScaleCriteria/ColorScaleCriteria 对象.html)** 集合，请使用 **[ColorScale](./../ColorScale/ColorScale 对象.html)** 对象的 **ColorScaleCriteria** 属性。

## 示例

```javascript
/*本示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后通过在ColorScaleCriteria集合中编制索引来设置单独的条件，从而指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 1
    ActiveSheet.Range("C2").Value2 = 2
    ActiveSheet.Range("C1:C2").AutoFill(Range("C1:C10"))
    Range("C1:C10").Select()

    // Create a two-color ColorScale object for the created sample data range
    let colorScale = Selection.FormatConditions.AddColorScale(2)

    // Set the minimum threshold to red and maximum threshold to blue
    colorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255)
}
```

```javascript
/*本示例设置活动工作表上单元格区域 A1:A10 的第二个色阶条件格式的所有阈值条件的颜色。*/
function test() {
    let criteria = ActiveSheet.Range("A1:A10").FormatConditions.Item(2).ColorScaleCriteria
    for (let i = 1; i <= criteria.Count; i++) {
        criteria.Item(i).FormatColor.ColorIndex = 6 + i
    }
}
```
