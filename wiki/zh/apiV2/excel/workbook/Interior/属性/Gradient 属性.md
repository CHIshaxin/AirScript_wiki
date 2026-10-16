---
title: "Gradient (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/Gradient%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BwnLAJwD
---

# Gradient (属性)

返回或设置所选内容的 **[Interior](./../../Interior/Interior 对象.html)** 对象的 **Gradient** 属性。只读。

## 说明

使用 LinearGradient 或 RectangularGradient。

## 示例

```javascript
/*本示例将 Sheet1 中 A1:E4 区域单元格内部设置为线性渐变填充，渐变填充角度为45度，然后添加 ColorStop 对象并设置其颜色为绿色。*/
function test() {
    let interior = Application.Worksheets.Item("Sheet1").Range("A1:E4").Interior
    interior.Pattern = xlPatternLinearGradient
    interior.Gradient.Degree = 45
    interior.Gradient.ColorStops.Add(1).Color = RGB(0, 255, 0)
}
```

```javascript
/*本示例显示活动工作表中 A1 单元格内部的线性渐变填充角度。*/
function test() {
    let interior = ActiveSheet.Range("A1").Interior
    console.log(`渐变角度为 ${interior.Gradient.Degree} 度`)
}
```
