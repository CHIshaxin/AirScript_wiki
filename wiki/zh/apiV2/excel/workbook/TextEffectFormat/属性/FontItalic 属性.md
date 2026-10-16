---
title: "FontItalic (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/FontItalic%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CQxzrIss
---

# FontItalic (属性)

如果指定艺术字中的字体是倾斜的，则返回 **msoTrue**。**MsoTriState** 类型，可读写。

## 说明

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** 不应用于该属性。 |
| **msoFalse** 指定的艺术字对象中的文字不是倾斜的。 |
| **msoTriStateMixed** 不应用于该属性。 |
| **msoTriStateToggle** 不应用于该属性。 |
| **msoTrue** 指定的艺术字对象中的文字是倾斜的。 |

## 示例

```javascript
/*本示例将第一张工作表中第四个形状的字体设置为斜体。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    worksheet.Shapes.Item(4).TextEffect.FontItalic = msoTrue
}
```

```javascript
/*此示例显示活动工作表的第二个艺术字是否为斜体。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.TextEffect.FontItalic == msoTrue)
}
```
