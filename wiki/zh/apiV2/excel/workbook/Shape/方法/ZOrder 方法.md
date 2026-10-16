---
title: "ZOrder (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/ZOrder%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 6DmN8pQ1
---

# ZOrder (方法)

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。

## 说明

使用 **ZOrderPosition** 属性可确定形状在 Z 顺序中的当前位置。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| ZOrderCmd | Kso.MsoZOrderCmd | 必填 | 指定根据其他形状将指定的形状移到什么位置。 |

## 示例

```javascript
/*此示例首先向活动工作表添加一个椭圆形，并将该形状置于底层。*/
function test() {
    let shape = ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)
    shape.ZOrder(msoSendToBack)
}
```

```javascript
/*本示例将第一张工作表中第一个形状置于顶层。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ZOrder(msoBringToFront)
}
```
