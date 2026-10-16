---
title: "ID 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/ID%20%E5%B1%9E%E6%80%A7.html
---
# ID (属性)

返回一个 Long 值，它代表指定对象的类型。

## 说明

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

## 示例

```javascript
/*本示例显示活动工作表中的第一个形状ID。*/
function test() {
    let shapes = ActiveSheet.Shapes
    console.log(shapes.Item(1).ID)
}
```

```javascript
/*本示例遍历活动工作表所有形状并显示它们的ID。*/
function test() {
    let shapes = ActiveSheet.Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        console.log(shapes.Item(i).ID)
    }
}
```
