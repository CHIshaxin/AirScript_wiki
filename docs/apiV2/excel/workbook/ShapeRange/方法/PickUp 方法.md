---
title: "PickUp 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/PickUp%20%E6%96%B9%E6%B3%95.html
---
# PickUp (方法)

复制指定形状的格式。使用 **Apply**方法可将复制的格式应用到其他形状。

## 示例

```javascript
/*本示例复制第一张工作表中新建的ShapeRange对象的格式，然后将复制的格式应用到第二个新建的ShapeRange对象上，删除第一个ShapeRange对象。*/
function test() {
    let shpRange1 = Application.Worksheets.Item(1).Shapes.Range([3])
    let shpRange2 = Application.Worksheets.Item(1).Shapes.Range([2, 1])
    shpRange1.PickUp()
    shpRange2.Apply()
    shpRange1.Delete()
}
```

```javascript
/*此示例复制活动工作表上新建的ShapeRange对象的格式，然后将复制的格式应用到活动工作表所有形状。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1])
    shpRange.PickUp()
    ActiveSheet.Shapes.SelectAll()
    Selection.ShapeRange.Apply()
}
```
