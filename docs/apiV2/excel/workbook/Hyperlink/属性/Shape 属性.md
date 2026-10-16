---
title: "Shape 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/Shape%20%E5%B1%9E%E6%80%A7.html
---
# Shape (属性)

返回一个 **[Shape](./../../Shape/Shape 对象.html)** 对象，它代表附加到指定超链接的形状。

## 示例

```javascript
/*本示例获取第一个超链接所在的形状并设置填充颜色为红色*/
function test() {
    let shape = Application.Worksheets.Item("Sheet1").Hyperlinks.Item(1).Shape
    shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```

```javascript
/*本示例获取第二个超链接所在的形状并显示其名称*/
function test() {
    let shape = Application.Worksheets.Item("Sheet1").Hyperlinks.Item(2).Shape
    console.log(`超链接所在形状的名称是 ${shape.Name}`)
}
```
