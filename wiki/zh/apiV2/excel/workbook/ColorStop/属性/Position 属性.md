---
title: "Position (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStop/%E5%B1%9E%E6%80%A7/Position%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BfsSj6X5
---

# Position (属性)

返回或设置 **[ColorStop](./../../ColorStop/ColorStop 对象.html)** 对象的位置。可读/写。

## 返回值

Double

## 示例

```javascript
/*本示例显示活动单元格的第二个ColorStop对象的位置。*/
function test() {
    let colorStop = ActiveCell.Interior.Gradient.ColorStops.Item(2)
    console.log(colorStop.Position)
}
```

```javascript
/*本示例将当前选定区域的第一个ColorStop对象的位置设置为0.8。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Selection.Interior.Gradient.ColorStops.Item(1)
    colorStop.Position = 0.8
}
```
