---
title: "HorizontalFlip (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/HorizontalFlip%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DKPwXB0b
---

# HorizontalFlip (属性)

如果指定的形状绕水平对称轴翻转，则为 **True**。**MsoTriState**，只读。

## 示例

```javascript
/*本示例将第一张工作表中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).HorizontalFlip) {
            shapes.Item(i).Flip(msoFlipHorizontal)
        }
        if (shapes.Item(i).VerticalFlip) {
            shapes.Item(i).Flip(msoFlipVertical)
        }
    }
}
```

```javascript
/*本示例显示活动工作表中第一个形状是否进行过水平翻转。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(1).HorizontalFlip) {
        console.log("已水平翻转")
    }else {
        console.log("未水平翻转")
    }
}
```
