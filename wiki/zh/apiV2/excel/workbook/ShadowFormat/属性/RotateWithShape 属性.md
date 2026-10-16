---
title: "RotateWithShape (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/RotateWithShape%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DKOlSkVR
---

# RotateWithShape (属性)

返回或设置一个 **MsoTriState** 类型的值，该值表示是否在旋转形状时旋转阴影。可读/写。

## 示例

```javascript
/*本示例显示第一张工作表中第二个形状是否在旋转形状时旋转阴影。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(2).Shadow
    console.log(shadow.RotateWithShape == msoTrue)
}
```

```javascript
/*本示例设置活动工作表中第三个形状在旋转形状时旋转阴影。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(3).Shadow
    shadow.RotateWithShape = msoTrue
}
```
