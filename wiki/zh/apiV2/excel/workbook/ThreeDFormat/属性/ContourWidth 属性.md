---
title: "ContourWidth (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/ContourWidth%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DFtKDD0w
---

# ContourWidth (属性)

返回或设置 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的轮廓线宽度。可读/写 **Single** 类型。

## 示例

```javascript
/*此示例将第一张工作表上第一个形状的三维轮廓宽度设置成10磅。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.ContourWidth = 10
}
```

```javascript
/*本示例显示活动工作表上第二个形状的三维轮廓宽度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.ContourWidth)
}
```
