---
title: "PictureEffects 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/PictureEffects%20%E5%B1%9E%E6%80%A7.html
---
# PictureEffects (属性)

返回一个对象，该对象表示指定的填充格式的图片或纹理填充。只读。

## 说明

可以在图表中各种元素（形状）的格式设置中指定图片或纹理填充。例如，可使用\*\*“设置数据系列格式”**对话框将**“列”**图表中各列的格式设置为图片或纹理填充。在此情况下，**PictureEffects** 属性返回的 **PictureEffects** 集合对应于**“设置数据系列格式”**对话框的**“填充”**类别中**“图片或纹理填充”\*\*选项的关联设置。

## 示例

```javascript
/*本示例在第一张工作表中第一个形状的填充中将水泥效果添加到效果链。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    fillFormat.PictureEffects.Insert(msoEffectCement, 1)
}
```

```javascript
/*本示例显示活动工作表中第二个形状中效果链的图片效果数量。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    console.log(fillFormat.PictureEffects.Count)
}
```
