---
title: "Connector (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Connector%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BwXNz0Xf
---

# Connector (属性)

如果指定的形状是连接符，则此属性为 **True**。**MsoTriState** 类型，只读。

## 示例

```javascript
/*本示例删除了活动工作表中所有的连接符。*/
function test() {
    let shapes = ActiveSheet.Shapes
    for (let i = shapes.Count; i >= 1; i--) {
        if (shapes.Item(i).Connector) {
            shapes.Item(i).Delete()
        }
    }
}
```

```javascript
/*本示例判断第一张工作表的第三个形状是否是连接符。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(3)
    if (shape.Connector) {
        console.log("是连接符")
    }else {
        console.log("不是连接符")
    }
}
```
