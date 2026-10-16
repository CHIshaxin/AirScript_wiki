---
title: "Width (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C8fVEVgI
---

# Width (属性)

返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*本示例显示第一张工作表中第一个形状的宽度。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    console.log(shapes.Item(1).Width)
}
```

```javascript
/*本示例将活动工作表中第二个形状的宽度减小一半。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.Item(2).Width *= 0.5
}
```
