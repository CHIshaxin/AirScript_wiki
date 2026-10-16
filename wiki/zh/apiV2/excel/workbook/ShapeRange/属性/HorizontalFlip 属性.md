---
title: "HorizontalFlip (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/HorizontalFlip%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B2CoJtQr
---

# HorizontalFlip (属性)

如果指定的形状绕水平对称轴翻转，则为 **True**。**MsoTriState**，只读。

## 示例

```javascript
/*本示例判断如果第一张工作表中新建的ShapeRange对象进行过水平翻转，则还原至初始状态。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    if (shpRange.HorizontalFlip) {
        shpRange.Flip(msoFlipHorizontal)
    }
}
```

```javascript
/*本示例判断活动工作表的所有形状是否进行过水平翻转。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    if (Selection.ShapeRange.HorizontalFlip) {
        console.log("已水平翻转")
    } else {
        console.log("未水平翻转")
    }
}
```
