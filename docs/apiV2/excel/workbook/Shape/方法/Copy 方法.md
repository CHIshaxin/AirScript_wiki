---
title: "Copy 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Copy%20%E6%96%B9%E6%B3%95.html
---
# Copy (方法)

将对象复制到剪贴板。

## 示例

```javascript
/*本示例将第一张工作表中的第二个形状复制到剪贴板。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(2).Copy()
}
```

```javascript
/*本示例判断如果活动工作表中第三个形状为平行四边形，则复制到剪贴板。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(3).AutoShapeType == msoShapeParallelogram) {
        shapes.Item(3).Copy()
    }
}
```
