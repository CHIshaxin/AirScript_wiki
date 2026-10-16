---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B4qKbVwo
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象删除。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Delete()
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象删除。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range(["正五边形 1", "椭圆 3"])
    shpRange.Delete()
}
```
