---
title: "Left (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BBqm8eeo
---

# Left (属性)

返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例显示活动工作表中第一个形状的左边缘到活动工作表左边缘的距离。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    console.log(shape.Left)
}
```

```javascript
/*本示例将第一张工作表中第二个形状的左边缘向左移动30。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(2)
    shape.Left = shape.Left - 30
}
```
