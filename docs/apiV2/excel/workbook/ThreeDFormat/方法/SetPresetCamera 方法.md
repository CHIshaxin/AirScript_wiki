---
title: "SetPresetCamera 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E6%96%B9%E6%B3%95/SetPresetCamera%20%E6%96%B9%E6%B3%95.html
---
# SetPresetCamera (方法)

为指定的 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象设置相机。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| PresetCamera | Kso.MsoPresetCamera | 必填 | 指定预设相机。 |

## 示例

```javascript
/*本示例将第一张工作表中的第一个形状ThreeDFormat对象的预设相机类型设置成等长底部朝下。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let threeDFormat = worksheet.Shapes.Item(1).ThreeD
    threeDFormat.SetPresetCamera(msoCameraIsometricBottomDown)
}
```

```javascript
/*本示例将活动工作表中第三个形状ThreeDFormat对象的预设相机类型设置成等轴左上。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    shape.ThreeD.Visible = msoTrue
    shape.ThreeD.SetPresetCamera(msoCameraIsometricLeftUp)
}
```
