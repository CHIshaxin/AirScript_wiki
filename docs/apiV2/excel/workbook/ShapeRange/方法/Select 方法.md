---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Replace | any | 可选 | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 示例

```javascript
/*本示例选中第一张工作表中新建的ShapeRange1对象，并追加新建的ShapeRange2对象。*/
function test() {
    let shpRange1 = Application.Worksheets.Item(1).Shapes.Range([1])
    shpRange1.Select()
    let shpRange2 = Application.Worksheets.Item(1).Shapes.Range([3])
    shpRange2.Select(false)
}
```

```javascript
/*本示例选中活动工作表中新建的ShapeRange对象。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1, 2, 3])
    shpRange.Select()
}
```
