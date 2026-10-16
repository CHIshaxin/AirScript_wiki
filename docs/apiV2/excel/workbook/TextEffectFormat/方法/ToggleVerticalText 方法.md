---
title: "ToggleVerticalText 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E6%96%B9%E6%B3%95/ToggleVerticalText%20%E6%96%B9%E6%B3%95.html
---
# ToggleVerticalText (方法)

将指定艺术字中的文字排列方向从水平切换为垂直，反之亦然。

## 说明

使用 **ToggleVerticalText** 方法可以交换代表艺术字的 **[Shape](./../../Shape/Shape 对象.html)** 对象的 **Width** 和 **Height** 属性的值，但不改变 **Left** 和 **Top** 属性。

**[Shape](./../../Shape/Shape 对象.html)** 对象的 **Flip** 方法和 **Rotation** 属性以及 **[TextEffectFormat](./../../TextEffectFormat/TextEffectFormat 对象.html)** 对象的 **RotatedChars** 属性和 **ToggleVerticalText** 方法都会影响代表艺术字的 **[Shape](./../../Shape/Shape 对象.html)** 对象中字符方向和文本排列方式。可能必须经过试验才能找到组合使用这些属性和方法的最佳途径以达到预期效果。

## 示例

```javascript
/*本示例将包含文本 Test 的艺术字添加到第一张工作表，并从水平文本流 (指定艺术字样式 的默认值 msoTextEffect1) 切换到垂直文本流。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let newWordArt = worksheet.Shapes.AddTextEffect(msoTextEffect1, "Test", "Arial Black", 36, false, false, 100, 100)
    newWordArt.TextEffect.ToggleVerticalText()
}
```

```javascript
/*此示例将活动工作表的第四个艺术字中的文字排列从水平切换为垂直。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(4)
    shape.TextEffect.ToggleVerticalText()
}
```
