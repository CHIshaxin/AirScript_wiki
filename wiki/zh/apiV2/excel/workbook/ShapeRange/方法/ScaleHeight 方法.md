---
title: "ScaleHeight (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/ScaleHeight%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D36kTAQz
---

# ScaleHeight (方法)

按指定的比例调整形状的高度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整高度。

## 说明

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue**。不应用于此属性。 |
| **msoFalse**。相对于形状的当前尺寸来调整形状的大小。 |
| **msoTriStateMixed**。不应用于此属性。 |
| **msoTriStateToggle**。不应用于此属性。 |
| **msoTrue**。相对于形状的初始尺寸来调整形状的大小。 |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Factor | float | 必填 | 指定形状调整后的高度与当前或原始高度的比例。例如，若要将一个矩形放大百分之五十，请将此参数设为 1.5。 |
| RelativeToOriginalSize | Kso.MsoTriState | 必填 | 如果为 msoTrue，则相对于形状的原有尺寸来调整高度。如果该值为 msoFalse，则相对于形状的当前尺寸来调整高度。仅当指定的形状是图片或 OLE 对象时，才能将此参数指定为 msoTrue。 |
| Scale | any | 可选 | MsoScaleFrom 的常量之一，它指定调整形状大小时，该形状哪一部分的位置将保持不变。 |

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象放大至原高度和宽度的 175%。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([2, 1])
    shpRange.ScaleHeight(1.75, msoTrue)
    shpRange.ScaleWidth(1.75, msoTrue)
}
```

```javascript
/*本示例将活动工作表的所有形状缩小至当前尺寸的50%。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.ScaleHeight(0.5, msoFalse, msoScaleFromTopLeft)
}
```
