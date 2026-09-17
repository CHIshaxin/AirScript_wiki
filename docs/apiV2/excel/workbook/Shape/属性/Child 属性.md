---
title: "Child 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Child%20%E5%B1%9E%E6%80%A7.html
---
# Child (属性)

如果指定的形状是子形状，或者如果形状区域中的所有形状都是同一个父形状的子形状，则返回 **msoTrue**。**MsoTriState** 类型，只读。

## 说明

| **MsoTriState** 可为以下 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue**。不应用于此属性。 |
| 如果选择的形状不是子形状，则为 **msoFalse**。 |
| 如果只有一些选择的形状是子形状，则为 **msoTriStateMixed**。 |
| **msoTriStateToggle**。不应用于此属性。 |
| 如果选择的形状是子形状，则为 **msoTrue**。 |

## 示例

```javascript
/*本示例判断活动工作表第一个形状的第一项单个图形是否为子形状并显示结果。*/
function test() {
    let shpGroupItem = ActiveSheet.Shapes.Item(1).GroupItems.Item(1)
    if (shpGroupItem.Child == msoTrue) {
        console.log("第一个形状是子形状")
    } else if (shpGroupItem.Child == msoTriStateMixed) {
        console.log("其中只有一部分形状是子形状")
    } else {
        console.log("第一个形状不是子形状")
    }
}
```

```javascript
/*本示例判断第一张工作表第一个形状的第二项单个图形如果是子形状，则用指定的颜色填充该子形状。*/
function test() {
    let shpGroupItem = Worksheets.Item(1).Shapes.Item(1).GroupItems.Item(2)
    if (shpGroupItem.Child == msoTrue) {
        shpGroupItem.Fill.ForeColor.RGB = RGB(100, 0, 200)
    }
}
```
