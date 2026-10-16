---
title: "ZOrder (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/ZOrder%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DeDSzcvC
---

# ZOrder (方法)

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。

## 说明

| **MsoZOrderCmd** 可为以下 **MsoZOrderCmd** 常量之一。 |
| --- |
| **msoBringForward** |
| **msoBringInFrontOfText**。仅用于 WPS。 |
| **msoBringToFront** |
| **msoSendBackward** |
| **msoSendBehindText**。仅用于 WPS。 |
| **msoSendToBack** |

使用 **ZOrderPosition** 属性可确定形状在 z-次序中的当前位置。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| ZOrderCmd | Kso.MsoZOrderCmd | 必填 | 指定根据其他形状将指定的形状移到什么位置。 |

## 示例

```javascript
/*本示例先在第一张工作表中创建shpRange对象，并将该对象置于底层。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([3, 2])
    shpRange.ZOrder(msoSendToBack)
}
```

```javascript
/*本示例先在活动工作表中创建shpRange对象，并将该对象置于顶层。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1, 2])
    shpRange.ZOrder(msoBringToFront)
}
```
