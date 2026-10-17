---
title: "BevelTopType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/BevelTopType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: FBBadeYR
---

# BevelTopType (属性)

返回或设置 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的顶部凹凸效果类型。可读/写 **MsoBevelType** 类型。

## 说明

| MsoBevelType 可以是下列 **MsoBevelType** 常量之一。 |
| --- |
| **msoBevelAngle** |
| **msoBevelArtDeco** |
| **msoBevelCircle** |
| **msoBevelConvex** |
| **msoBevelCoolSlant** |
| **msoBevelCross** |
| **msoBevelDivot** |
| **msoBevelHardEdge** |
| **msoBevelNone** |
| **msoBevelRelaxedInset** |
| **msoBevelRiblet** |
| **msoBevelSlope** |
| **msoBevelSoftRound** |
| **msoBevelTypeMixed** |

## 示例

```javascript
/*此示例将第一张工作表的第三个形状顶部凹凸效果类型设置为msoBevelCircle。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(3)
    shape.ThreeD.Visible = msoTrue
    shape.ThreeD.BevelTopType = msoBevelCircle
}
```

```javascript
/*本示例显示活动工作表中第二个形状顶部凹凸效果类型是否为msoBevelRelaxedInset。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.BevelTopType == msoBevelRelaxedInset)
}
```
