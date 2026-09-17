---
title: "Height 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
---
# Height (属性)

返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*本示例先在第一张工作表创建ShapeRange对象，再将其高度设置为200。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range(["矩形 1", "椭圆 2"])
    shpRange.Height = 200
}
```

```javascript
/*本示例显示新建的ShapeRange对象中第二个形状的高度。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range(["矩形 1", "椭圆 2"])
    console.log(shpRange.Item(2).Height)
}
```
