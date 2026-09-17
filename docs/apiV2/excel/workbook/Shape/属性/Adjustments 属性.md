---
title: "Adjustments 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Adjustments%20%E5%B1%9E%E6%80%A7.html
---
# Adjustments (属性)

返回一个 **[Adjustments](./../../Adjustments/Adjustments 对象.html)** 对象，它包括指定形状的所有调整操作的调整值。应用于任何代表自选图形、艺术字或连接符的 **[Shape](./../../Shape/Shape 对象.html)** 对象。

## 示例

```javascript
/*本示例将活动工作表上第一个形状的第一个调整值设置为 0.25。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.Adjustments.Item(1, 0.25)
}
```

```javascript
/*本示例对活动工作表上第二个形状进行调整，并返回活动工作表上第二个形状的第二个调整值。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.Adjustments.Item(1, 1.25)
    shape.Adjustments.Item(2, 2)
    console.log(shape.Adjustments.Item(2))
}
```
