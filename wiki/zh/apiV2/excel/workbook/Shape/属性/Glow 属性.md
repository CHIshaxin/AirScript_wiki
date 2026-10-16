---
title: "Glow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Glow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D6tLZIBX
---

# Glow (属性)

返回一个包含形状发光格式属性的指定形状的 **GlowFormat** 对象。只读。

## 说明

发光效果为图形添加了色彩鲜明的彩色边缘。

## 示例

```javascript
/*本示例将活动工作表的第一个形状添加一个红色的发光效果，不透明度为0.5。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.Glow.Color.RGB = RGB(255, 0, 0)
    shape.Glow.Transparency = 0.5
}
```

```javascript
/*本示例将第一张工作表的第二个形状添加一个蓝色的发光效果，大小为10。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(2)
    shape.Glow.Color.RGB = RGB(0, 0, 255)
    shape.Glow.Size = 10
}
```
