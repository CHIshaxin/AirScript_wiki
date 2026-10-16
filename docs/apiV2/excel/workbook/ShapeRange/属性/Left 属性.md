---
title: "Left 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
---
# Left (属性)

返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象的左边缘设置为130。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Left = 130
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象的左边缘向右移动50。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([3])
    shpRange.Left += 50
}
```
