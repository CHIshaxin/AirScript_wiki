---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例将活动工作表中第一个形状的名称设置为"蓝色三角形"。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.Item(1).Name = "蓝色三角形"
}
```

```javascript
/*本示例将第一张工作表的第二个形状的高度增加一倍，并改名为"平行四边形"。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(2)
    shape.Height = shape.Height * 2
    shape.Name = "平行四边形"
}
```
