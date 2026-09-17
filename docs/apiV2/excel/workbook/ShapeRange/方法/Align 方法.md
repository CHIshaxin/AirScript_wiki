---
title: "Align 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Align%20%E6%96%B9%E6%B3%95.html
---
# Align (方法)

对齐指定形状区域中的形状。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| AlignCmd | Kso.MsoAlignCmd | 必填 | 指定某个指定形状范围中形状的对齐方式。 |
| RelativeTo | Kso.MsoTriState | 必填 | 不在 ET 中使用。必须为 False。 |

## 示例

```javascript
/*本示例将第一张工作表的所有形状的左边与该范围内最左边的形状的左边对齐。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Align(msoAlignLefts, false)
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象的右边与该范围内最右边的形状的右边对齐。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([2, 3])
    shpRange.Align(msoAlignRights, false)
}
```
