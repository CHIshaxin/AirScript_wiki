---
title: "Duplicate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Duplicate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CCB4fENY
---

# Duplicate (方法)

复制对象，并返回对新复制对象的引用。

## 示例

```javascript
/*本示例复制第一张工作表中新建的ShapeRange对象并设置填充格式的前景色。*/
function test() {
    let duplicate = Application.Worksheets.Item(1).Shapes.Range([1, 2]).Duplicate()
    duplicate.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```

```javascript
/*本示例复制活动工作表的所有形状。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Duplicate()
}
```
