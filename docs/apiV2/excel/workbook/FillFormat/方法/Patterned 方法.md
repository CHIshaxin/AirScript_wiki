---
title: "Patterned 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E6%96%B9%E6%B3%95/Patterned%20%E6%96%B9%E6%B3%95.html
---
# Patterned (方法)

将指定填充设为一种图案。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Pattern | Kso.MsoPatternType | 必填 | 图案的类型。 |

## 示例

```javascript
/*本示例在活动工作表中添加矩形，并将前景色设置为红色，填充图案设置为网格。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
    fillFormat.ForeColor.RGB = RGB(255, 0, 0)
    fillFormat.Patterned(msoPatternTrellis)
}
```

```javascript
/*本示例在第一张工作表中添加椭圆，并将前景色设置为蓝色，填充图案设置为草皮。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).Fill
    fillFormat.ForeColor.RGB = RGB(0, 0, 255)
    fillFormat.Patterned(msoPatternDivot)
}
```
