---
title: "ColorStops (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LinearGradient/%E5%B1%9E%E6%80%A7/ColorStops%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BM03XH2g
---

# ColorStops (属性)

返回 **[LinearGradient](./../../LinearGradient/LinearGradient 对象.html)** 对象的 **ColorStops**。只读。

## 返回值

ColorStops

## 示例

```javascript
/*下例将第一张工作表上区域 B1:B10 线性渐变的第一个颜色停止点的颜色设置为蓝色。*/
function test() {
    let linearGradient = Worksheets.Item(1).Range("B1:B10").Interior.Gradient
    let colorStop1 = linearGradient.ColorStops.Item(1)
    colorStop1.Color = RGB(0, 0, 255)
}
```

```javascript
/*下例清除活动工作表上区域 B1:B10 中线性渐变所有已经设置的颜色停止点。*/
function test() {
    let linearGradient = ActiveSheet.Range("B1:B10").Interior.Gradient
    linearGradient.ColorStops.Clear()
}
```
