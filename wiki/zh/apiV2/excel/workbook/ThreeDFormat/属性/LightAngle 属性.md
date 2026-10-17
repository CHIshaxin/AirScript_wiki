---
title: "LightAngle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/LightAngle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DKeBMJUW
---

# LightAngle (属性)

返回或设置对 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象设置的延伸灯光的角度。可读/写 **Single** 类型。

## 示例

```javascript
/*此示例将第一张工作表的第一个形状延伸灯光的角度设置为20度。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.LightAngle = 20
}
```

```javascript
/*本示例显示活动工作表中第二个形状延伸灯光的角度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.LightAngle)
}
```
