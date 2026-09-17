---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

返回一个代表形状类型的 **MsoShapeType** 值。只读。

## 示例

```javascript
/*本示例判断第一张工作表中第一个形状的类型是否为图片。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    console.log(shapes.Item(1).Type == msoPicture)
}
```

```javascript
/*本示例判断活动工作表中第二个形状是否为图表并显示结果。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(2).Type == msoChart) {
        console.log("This is a Chart")
    } else {
        console.log("This is not a Chart")
    }
}
```
