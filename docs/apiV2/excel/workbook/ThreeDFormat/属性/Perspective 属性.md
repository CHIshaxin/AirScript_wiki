---
title: "Perspective 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/Perspective%20%E5%B1%9E%E6%80%A7.html
---
# Perspective (属性)

返回或设置一个 **MsoTriState** 值，该值确定是否在透视中显示延伸。

## 说明

如果延伸是平行投影或正交投影，则返回 **msoFalse** ，也就是说，墙壁不会向消失点变窄。

如果延伸显示在透视中，则返回 **msoTrue** ，即延伸的墙壁向消失点变窄。

## 示例

```javascript
/*此示例设置第一张工作表的第一个形状在透视中显示延伸。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.Perspective = msoTrue
}
```

```javascript
/*本示例显示活动工作表中第三个形状是否在透视中显示延伸。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ThreeD.Perspective == msoTrue)
}
```
