---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CB8lF6SS
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例将第一张工作表中的所有椭圆删除。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).AutoShapeType == msoShapeOval) {
            shapes.Item(i).Delete()
        }
    }
}
```

```javascript
/*本示例判断如果活动工作表中第二个形状为矩形，则删除。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(2).AutoShapeType == msoShapeRectangle) {
        shapes.Item(2).Delete()
    }
}
```
