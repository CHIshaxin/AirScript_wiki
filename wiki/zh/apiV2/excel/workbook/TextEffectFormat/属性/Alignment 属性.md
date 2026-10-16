---
title: "Alignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/Alignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DegduKVr
---

# Alignment (属性)

返回或设置一个 **MsoTextEffectAlignment** 值，它代表艺术字的对齐方式。

## 示例

```javascript
/*此示例向第一张工作表中添加艺术字对象，然后将该对象设置右对齐。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shape = shapes.AddTextEffect(msoTextEffect1, "Test Text", "Palatino", 54, true, false, 100, 50)
    shape.TextEffect.Alignment = msoTextEffectAlignmentRight
}
```

```javascript
/*此示例显示活动工作表的第一个艺术字的对齐方式是否为居中对齐。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    console.log(shape.TextEffect.Alignment == msoTextEffectAlignmentCentered)
}
```
