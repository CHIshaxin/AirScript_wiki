---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 说明

对象的文本名称就是 **Name** 属性的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

包含在集合中的一个 **[Shape](./../../Shape/Shape 对象.html)** 对象。

## 示例

```javascript
/*本示例对 Shapes 集合中第二个形状的 OnAction 属性进行设置。如果变量 ss 代表的不是 Shapes 对象，则本示例无效。*/
function test() {
    let ss = Application.Worksheets.Item(1).Shapes
    ss.Item(2).OnAction = "ShapeAction"
}
```

```javascript
/*本示例将活动工作表中不包含图表的形状删除。*/
function test() {
    let shapes = ActiveSheet.Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).HasChart) {
            shapes.Item(i).Delete()
        }
    }
}
```
