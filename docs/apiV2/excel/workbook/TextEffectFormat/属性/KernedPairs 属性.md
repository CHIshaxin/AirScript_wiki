---
title: "KernedPairs 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/KernedPairs%20%E5%B1%9E%E6%80%A7.html
---
# KernedPairs (属性)

如果指定艺术字中的字符对是自动缩紧的，则该属性值为 **msoTrue**。**MsoTriState** 类型，可读写。

## 说明

| **MsoTriState**可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** |
| **msoFalse** |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue**指定艺术字中的字符对自动紧缩。 |

## 示例

```javascript
/*本示例打开第一张工作表中第三个形状（如果该形状为艺术字）的字符对自动紧缩。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let shape = worksheet.Shapes.Item(3)
    if (shape.Type == msoTextEffect) {
        shape.TextEffect.KernedPairs = msoTrue
    }
}
```

```javascript
/*此示例显示活动工作表的第二个艺术字的字符对是否打开自动紧缩。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.TextEffect.KernedPairs == msoTrue)
}
```
