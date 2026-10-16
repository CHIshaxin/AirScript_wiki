---
title: "OnAction (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/OnAction%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BvwTeE7e
---

# OnAction (属性)

返回或设置单击指定对象时运行的宏的名称。**String** 类型，可读写。

## 说明

为菜单项设置该属性，该属性可取代为菜单项和指定宏设定的自定义帮助信息。

## 示例

```javascript
/*本示例在选择第一张工作表的第一个形状时运行 ShapeClick 过程。*/
function test() {
    Worksheets(1).Shapes(1).OnAction = "ShapeClick"
}
```

```javascript
/*本示例在选择活动工作表的第二个形状时运行 ShapeAction 过程。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.OnAction = "ShapeAction"
}
```
