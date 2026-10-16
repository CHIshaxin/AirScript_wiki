---
title: "ShadowFormat (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/ShadowFormat%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: DPao8szt
---

# ShadowFormat (对象)

代表形状的阴影格式。

## 说明

使用 **Shadow**属性可返回一个 **[ShadowFormat](./../ShadowFormat/ShadowFormat 对象.html)** 对象。

## 示例

```javascript
/*下例给第一张工作表添加带阴影的矩形。半透明蓝色阴影在矩形右侧偏移 5 磅，在矩形上方偏移 3 磅。*/
function test() {
    let shadow = Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200).Shadow
    shadow.ForeColor.RGB = RGB(0, 0, 128)
    shadow.Type = msoShadow17
    shadow.OffsetX = 5
    shadow.OffsetY = -3
    shadow.Transparency = 0.5
    shadow.Visible = true
}
```

```javascript
/*本示例显示活动工作表中第二个形状阴影的大小。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(2).Shadow
    console.log(shadow.Size)
}
```
