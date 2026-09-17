---
title: "Child 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Child%20%E5%B1%9E%E6%80%A7.html
---
# Child (属性)

如果指定的形状是子形状，或者如果形状区域中的所有形状都是同一个父形状的子形状，则返回 **msoTrue**。**MsoTriState** 类型，只读。

## 说明

| **MsoTriState** 可为以下 **MsoTriState** 常量之一: |
| --- |
| **msoCTrue**。不应用于此属性。 |
| 如果选择的形状不是子形状，则为 **msoFalse**。 |
| 如果只有一些选择的形状是子形状，则为 **msoTriStateMixed**。 |
| **msoTriStateToggle**。不应用于此属性。 |
| 如果选择的形状是子形状，则为 **msoTrue**。 |

## 示例

```javascript
/*本示例选择第一张工作表中部分形状，并且如果选择的形状是子形状，则用指定的颜色填充该形状。*/
function test() {
    Worksheets.Item(1).Shapes.Item(1).GroupItems.Item(1).Select()
    Worksheets.Item(1).Shapes.Item(1).GroupItems.Item(2).Select(false)
    if (Selection.ShapeRange.Child == msoTrue) {
        Selection.ShapeRange.Fill.ForeColor.RGB = (100, 0, 200)
    } else {
        console.log("有部分形状不是子形状")
    }
}
```

```javascript
/*本示例判断所选形状是否都是同一个父形状的子形状，并显示。本示例假定所选形状都是子形状。*/
function test() {
    ActiveSheet.Shapes.Item(2).GroupItems.Item(1).Select()
    ActiveSheet.Shapes.Item(2).GroupItems.Item(2).Select(false)
    if (Selection.ShapeRange.Child == msoTrue) {
        console.log("所选形状都是同一个父形状的子形状")
    } else {
        console.log("所选形状不是同一个父形状的子形状")
    }
}
```
