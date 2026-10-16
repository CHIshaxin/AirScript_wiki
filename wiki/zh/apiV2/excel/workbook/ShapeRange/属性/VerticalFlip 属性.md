---
title: "VerticalFlip (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/VerticalFlip%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D3ce4yv0
---

# VerticalFlip (属性)

如果指定的形状绕垂直坐标轴翻转，则此属性为 **True**。**MsoTriState** 类型，只读。

## 示例

```javascript
/*本示例先在第一张工作表创建shpRange对象，将对象中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    if (shpRange.HorizontalFlip) {
        shpRange.Flip(msoFlipHorizontal)
    }
    if (shpRange.VerticalFlip) {
        shpRange.Flip(msoFlipVertical)
    }
}
```

```javascript
/*本示例判断活动工作表中创建shpRange对象是否围绕垂直轴进行翻转，并显示结果。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1])
    if (shpRange.VerticalFlip) {
        console.log("已围绕垂直轴进行翻转")
    } else {
        console.log("未围绕垂直轴进行翻转")
    }
}
```
