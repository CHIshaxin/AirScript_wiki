---
title: "NoTextRotation 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/NoTextRotation%20%E5%B1%9E%E6%80%A7.html
---
# NoTextRotation (属性)

返回或设置在指定的对象旋转时，文本是否保持扁平。可读/写

## 说明

如果文本框旋转时，文本仍保持扁平，则为 **msoTrue** (-1)；否则为 **msoFalse** (0)。

## 示例

```javascript
/*本示例将活动工作表中第一个文本框的文本设置为随文本框旋转。*/
function test() {
    let textFrame2 = Application.ActiveSheet.Shapes.Item(1).TextFrame2
    textFrame2.NoTextRotation = msoFalse
}
```

```javascript
/*本示例显示第一张工作表中第二个形状的文本框旋转时，文本是否保持水平。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.Item(2).TextFrame2
    console.log(textFrame2.NoTextRotation == msoTrue)
}
```
