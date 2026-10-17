---
title: "BevelBottomType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/BevelBottomType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DNa-tz0u
---

# BevelBottomType (属性)

返回或设置 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的底部凹凸效果类型。可读/写 **MsoBevelType** 类型。

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
/*此示例将第一张工作表的第一个形状底部凹凸效果类型设置为msoBevelArtDeco。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.Visible = msoTrue
    shape.ThreeD.BevelBottomType = msoBevelArtDeco
}
```

```javascript
/*本示例显示活动工作表中第二个形状底部凹凸效果类型是否为msoBevelSlope。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.BevelBottomType == msoBevelSlope)
}
```
