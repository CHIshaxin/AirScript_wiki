---
title: "RotateWithObject (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/RotateWithObject%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CZmR1WKA
---

# RotateWithObject (属性)

返回或设置是否随对象旋转。可读/写 **MsoTriState** 类型。

## 说明

| MsoTriState 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** |
| **msoFalse** |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue** |

## 示例

```javascript
/*本示例设置第一张工作表中第二个形状填充不与形状一起旋转。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(2).Fill
    fillFormat.RotateWithObject = msoFalse
}
```

```javascript
/*本示例显示活动工作表中第一个形状填充是否与形状一起旋转。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(1).Fill
    console.log(fillFormat.RotateWithObject == msoTrue)
}
```
