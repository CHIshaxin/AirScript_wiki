---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/GroupShapes/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DCLyqYgI
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

包含在集合中的一个 Shape 对象。

## 示例

```javascript
/*本示例对某个形状区域中第二个形状的 OnAction 属性进行设置。如果变量 sr 不代表 ShapeRange 对象，则本示例无效。*/
function test(){
let sr = Application.Worksheets.Item(1).Shapes
sr.Item(2).OnAction = "ShapeAction"
}
```
