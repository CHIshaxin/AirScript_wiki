---
title: "RotatedChars 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/RotatedChars%20%E5%B1%9E%E6%80%A7.html
---
# RotatedChars (属性)

如果指定艺术字对象中的字符相对于该对象旋转了 90 度，则该值为 **msoTrue**。如果指定艺术字对象中的字符相对于该对象保持原有方向，则该值为 **msoFalse**。**MsoTriState** 类型，可读/写。

## 说明

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** |
| **msoFalse**指定艺术字中的字符保留原有的相对于边框形状的方向。 |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue**将指定艺术字中的字符相对于艺术字边框形状旋转 90 度。 |

如果艺术字中含有水平文字，将 **RotatedChars**属性设置为 **msoTrue** 可将该字符逆时针旋转 90 度。如果艺术字中含有垂直文字，将 **RotatedChars** 属性设置为 **msoFalse** 可将该字符顺时针旋转 90 度。使用 **ToggleVerticalText** 方法在水平文字流和垂直文字流之间切换。

**[Shape](./../../Shape/Shape 对象.html)** 对象的 **Flip** 方法和 **Rotation** 属性以及 **[TextEffectFormat](./../../TextEffectFormat/TextEffectFormat 对象.html)** 对象的 **RotatedChars** 属性和 **ToggleVerticalText** 方法会对表示艺术字的 **[Shape](./../../Shape/Shape 对象.html)** 对象中的字符的方向和文字排列的方向产生影响。可能必须经过试验才能找到组合使用这些属性和方法的最佳途径以达到预期效果。

## 示例

```javascript
/*本示例将包含文本“Test”的艺术字添加到第一张工作表，并将字符逆时针旋转 90 度。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let newWordArt = worksheet.Shapes.AddTextEffect(msoTextEffect1, "Test", "Arial Black", 36, false, false, 10, 10)
    newWordArt.TextEffect.RotatedChars = msoTrue
}
```

```javascript
/*此示例显示活动工作表的第二个艺术字的字符是否相对于艺术字的边框形状旋转 90 度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.TextEffect.RotatedChars == msoTrue)
}
```
