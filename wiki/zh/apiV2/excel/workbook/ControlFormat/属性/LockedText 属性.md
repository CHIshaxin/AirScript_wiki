---
title: "LockedText (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/LockedText%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: xzCdOGlv
---

# LockedText (属性)

当工作簿受保护时，指定对象中的文字是否被锁定以免发生更改，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将第一张工作表的第一个形状（组合框）的文字锁定。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ControlFormat.LockedText = true
}
```

```javascript
/*本示例显示活动工作表的第二个形状（列表框）的文字是否被锁定。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ControlFormat.LockedText)
}
```
