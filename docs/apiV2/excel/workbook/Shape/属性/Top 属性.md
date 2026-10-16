---
title: "Top 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
---
# Top (属性)

返回或设置一个 **Single** 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例设置活动工作表第一个形状上边缘到工作表上边缘的距离为150磅。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.Top = 150
}
```

```javascript
/*本示例显示第一张工作表中的第二个形状上边缘到工作表上边缘的距离。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(2)
    console.log(shape.Top)
}
```
