---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bj8-pB9L
---

# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例先在第一张工作表创建shpRange对象，再将其名称设置为“ShapeRange1”。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1])
    shpRange.Name = "ShapeRange1"
}
```

```javascript
/*本示例判断如果活动工作表新建的shpRange对象的名称是“圆角矩形2”，则删除该对象。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([2])
    if (shpRange.Name == "圆角矩形2") {
        shpRange.Delete()
    }
}
```
