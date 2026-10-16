---
title: "Width (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CVnB8_mR
---

# Width (属性)

返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*本示例显示第一张工作表中新建的ShapeRange对象的宽度。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 3])
    console.log(shpRange.Width)
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象的宽度设置为100磅。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1, 2])
    shpRange.Width = 100
}
```
