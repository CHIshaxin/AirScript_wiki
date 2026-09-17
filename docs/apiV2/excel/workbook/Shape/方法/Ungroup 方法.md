---
title: "Ungroup 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Ungroup%20%E6%96%B9%E6%B3%95.html
---
# Ungroup (方法)

取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。

## 说明

由于一组形状被视为单个对象，因此对形状进行组合或取消组合操作时，将会更改 **[Shapes](./../../Shapes/Shapes 对象.html)** 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

## 返回值

一个 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象，它代表取消组合的形状。

## 示例

```javascript
/*此示例将第一张工作表中所有的形状取消组合，若有非组合的形状会导致报错。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        shapes.Item(i).Ungroup()
    }
}
```

```javascript
/*此示例将活动工作表中所有的组合形状取消组合。*/
function test() {
    let shapes = ActiveSheet.Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).Type == msoGroup) {
            shapes.Item(i).Ungroup()
        }
    }
}
```
