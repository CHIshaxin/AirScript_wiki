---
title: "Adjustments 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Adjustments%20%E5%B1%9E%E6%80%A7.html
---
# Adjustments (属性)

返回一个 **[Adjustments](./../../Adjustments/Adjustments 对象.html)** 对象，它包括指定形状的所有调整操作的调整值。应用于任何代表自选图形、艺术字或连接符的 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象。

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象的第一个调整值设置为 0.5。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1])
    shpRange.Adjustments.Item(1, 0.5)
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象的第二个调整值设置为 0.8。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([2])
    shpRange.Adjustments.Item(2, 0.8)
}
```
