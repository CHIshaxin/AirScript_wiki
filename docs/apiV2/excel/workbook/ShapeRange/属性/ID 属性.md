---
title: "ID 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/ID%20%E5%B1%9E%E6%80%A7.html
---
# ID (属性)

返回一个 Long 值，它代表指定对象的类型。

## 说明

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

## 示例

```javascript
/*本示例显示第一张工作表中新建的ShapeRange对象的ID。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1])
    console.log(shpRange.ID)
}
```

```javascript
/*本示例显示活动工作表中新建的ShapeRange对象的ID。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range(["等腰三角形 2"])
    console.log(shpRange.ID)
}
```
