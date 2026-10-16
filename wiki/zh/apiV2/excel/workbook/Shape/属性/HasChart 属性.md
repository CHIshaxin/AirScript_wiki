---
title: "HasChart (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/HasChart%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cw1YIANH
---

# HasChart (属性)

返回形状是否包含图表。**MsoTriState** 类型，只读。

## 示例

```javascript
/*本示例显示活动工作表中第一个形状是否包含图表。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(1).HasChart) {
        console.log("包含图表")
    }else {
        console.log("不包含图表")
    }
}
```

```javascript
/*本示例将第一张工作表中包含图表的形状删除。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).HasChart) {
            shapes.Item(i).Delete()
        }
    }
}
```
