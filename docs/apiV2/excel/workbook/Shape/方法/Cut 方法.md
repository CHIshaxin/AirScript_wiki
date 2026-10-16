---
title: "Cut 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Cut%20%E6%96%B9%E6%B3%95.html
---
# Cut (方法)

将对象剪切到剪贴板。

## 示例

```javascript
/*本示例将第一张工作表中的第一个形状剪切到剪贴板。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).Cut()
}
```

```javascript
/*本示例判断如果活动工作表中第二个形状为六角星，则剪切到剪贴板。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(2).AutoShapeType == msoShape6pointStar) {
        shapes.Item(2).Cut()
    }
}
```
