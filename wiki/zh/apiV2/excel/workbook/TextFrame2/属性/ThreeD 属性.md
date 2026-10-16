---
title: "ThreeD (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/ThreeD%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CslGBDoD
---

# ThreeD (属性)

返回一个 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象，该对象包含指定文本的三维效果格式属性。只读。

## 示例

```javascript
/*本示例将活动工作表中第二个形状的文本框内的文本突出的深度设置为10磅。*/
function test() {
    let textFrame2 = ActiveSheet.Shapes.Item(2).TextFrame2
    textFrame2.ThreeD.Depth = 10
}
```

```javascript
/*本示例显示第一张工作表中第一个形状的文本框三维格式是否可见。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.Item(1).TextFrame2
    console.log(textFrame2.ThreeD.Visible == msoTrue)
}
```
