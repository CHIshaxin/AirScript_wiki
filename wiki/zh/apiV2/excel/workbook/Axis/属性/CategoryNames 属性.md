---
title: "CategoryNames (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/CategoryNames%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Ddyn7tFX
---

# CategoryNames (属性)

以文本数组形式返回或设置指定坐标轴中所有分类的名称。该属性既可设为一个数组，也可设为一个包含所有分类名称的 **[Range](./../../Range/Range 对象.html)** 对象。**Variant** 类型，可读写。

## 说明

该属性仅适用于分类轴。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一张图表的分类名设为该工作表区域 B1:B5 单元格中的值。*/
function test() {
    let sheet = Application.Worksheets.Item("Sheet1")
    let axis = sheet.ChartObjects(1).Chart.Axes(xlCategory)
    axis.CategoryNames = sheet.Range("B1:B5")
}
```

```javascript
/*本示例使用数组对图表工作表 Chart1 中图表的个别分类名进行设置。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.CategoryNames = ["1985", "1986", "1987", "1988", "1989"]
}
```
