---
title: "Apply (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Apply%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Bx3YsaZi
---

# Apply (方法)

应用通过 **PickUp**方法复制的指定形状格式。

## 示例

```javascript
/*本示例复制第一张工作表中新建的ShapeRange对象的格式，然后将复制的格式应用到第二个新建的ShapeRange对象上。*/
function test() {
    let shpRange1 = Application.Worksheets.Item(1).Shapes.Range([1])
    let shpRange2 = Application.Worksheets.Item(1).Shapes.Range([2, 3])
    shpRange1.PickUp()
    shpRange2.Apply()
}
```

```javascript
/*此示例复制活动工作表上新建的ShapeRange对象的格式，然后将复制的格式应用到活动工作表所有形状。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([3])
    shpRange.PickUp()
    ActiveSheet.Shapes.SelectAll()
    Selection.ShapeRange.Apply()
}
```
