---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CxFb2KLi
---

# Select (方法)

选择对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Replace | any | 可选 | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 示例

```javascript
/*本示例选中第一张工作表中的所有平行四边形。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).AutoShapeType == msoShapeParallelogram) {
            shapes.Item(i).Select(false)
        }
    }
}
```

```javascript
/*本示例将活动工作表的第一个和第三个形状选中。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.Item(1).Select()
    shapes.Item(3).Select(false)
}
```
