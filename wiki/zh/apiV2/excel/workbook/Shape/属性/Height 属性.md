---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C3e6eEva
---

# Height (属性)

返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*本示例显示第一张工作表中第一个形状的高度。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    console.log(shape.Height)
}
```

```javascript
/*本示例将活动工作表中第二个形状的高度缩小一半。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.Height = shape.Height * 0.5
}
```
