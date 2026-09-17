---
title: "IncrementTop 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/IncrementTop%20%E6%96%B9%E6%B3%95.html
---
# IncrementTop (方法)

以指定磅数垂直移动指定形状。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Increment | float | 必填 | 指定形状对象垂直移动的距离，以磅为单位。正值将形状下移，负值将形状上移。 |

## 示例

```javascript
/*本示例将活动工作表的第二个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.Fill.PresetTextured(msoTextureGranite)
    shape.IncrementLeft(70)
    shape.IncrementTop(-50)
    shape.IncrementRotation(30)
}
```

```javascript
/*本示例将第一张工作表的第一个形状向下移 190 磅，然后逆时针旋转 20 度。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(1)
    shape.IncrementTop(190)
    shape.IncrementRotation(-20)
}
```
