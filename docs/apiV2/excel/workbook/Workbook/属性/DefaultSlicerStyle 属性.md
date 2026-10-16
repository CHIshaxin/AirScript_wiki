---
title: "DefaultSlicerStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/DefaultSlicerStyle%20%E5%B1%9E%E6%80%A7.html
---
# DefaultSlicerStyle (属性)

从 **[TableStyles](./../../TableStyles/TableStyles 对象.html)** 对象中指定用作切片器默认样式的样式。可读/写。

## 返回值

Variant

## 示例

```javascript
/*本示例演示活动工作簿切片器默认样式*/
function test() {
    console.log(ActiveWorkbook.DefaultSlicerStyle)
}
```

```javascript
/*本示例演示向活动工作簿的 TableStyles 集合中添加 slicer_style1 样式，并指定 slicer_style1 为切片器默认样式*/
function test() {
    ActiveWorkbook.TableStyles.Add("slicer_style1")
    ActiveWorkbook.DefaultSlicerStyle = "slicer_style1"
}
```
