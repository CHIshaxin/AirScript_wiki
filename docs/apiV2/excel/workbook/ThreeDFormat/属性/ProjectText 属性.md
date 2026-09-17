---
title: "ProjectText 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/ProjectText%20%E5%B1%9E%E6%80%A7.html
---
# ProjectText (属性)

返回或设置指定 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象投影文本状态。可读/写 **MsoTriState** 类型。

## 说明

| MsoTriState 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** |
| **msoFalse** |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue** |

## 示例

```javascript
/*本示例将第一张工作表中的第一个形状ThreeDFormat对象投影文本状态设置成msoFalse。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let threeDFormat = worksheet.Shapes.Item(1).ThreeD
    threeDFormat.ProjectText = msoFalse
}
```

```javascript
/*本示例显示活动工作表中第三个形状ThreeDFormat对象投影文本状态是否为msoTrue。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ThreeD.ProjectText == msoTrue)
}
```
