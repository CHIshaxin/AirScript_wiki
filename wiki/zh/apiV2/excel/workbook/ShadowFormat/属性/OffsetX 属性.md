---
title: "OffsetX (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/OffsetX%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Db3lhJY1
---

# OffsetX (属性)

以磅为单位返回或设置指定形状的阴影的水平偏移量。正偏移值将阴影向右偏移，负偏移值将阴影向左偏移。可读写。**Single** 类型。

## 说明

使用 **IncrementOffsetX** 方法或 **IncrementOffsetY**方法可将阴影从当前位置水平或垂直轻微移动而无需指定其绝对位置。

## 示例

```javascript
/*本示例为第一张工作表中的形状 3 设置阴影的格式类型、水平偏移量和垂直偏移量。阴影相对形状向右偏移 5 磅，向上偏移 3 磅。如果形状没有阴影，本示例为其添加一个阴影。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow
    shadow.Visible = true
    shadow.Type = msoShadow17
    shadow.OffsetX = 5
    shadow.OffsetY = -3
}
```

```javascript
/*本示例显示活动工作表中第二个形状阴影的水平偏移量。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(2).Shadow
    console.log(shadow.OffsetX)
}
```
