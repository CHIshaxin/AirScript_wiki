---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: ByYqMUr3
---

# Orientation (属性)

返回或设置一个值，该值代表文本框方向。可读/写 **MsoTextOrientation** 类型。

## 说明

此属性值可设置为从 -90 到 90 度的整数值，或者以下 **MsoTextOrientation** 常量之一。

## 示例

```javascript
/*本示例显示活动工作表中第二个形状的文本框方向是否为水平。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame2
    console.log(textFrame.Orientation == msoTextOrientationHorizontal)
}
```

```javascript
/*本示例将第一张工作表中第一个形状的文本框方向设置为向下。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame2.Orientation = msoTextOrientationDownward
}
```
