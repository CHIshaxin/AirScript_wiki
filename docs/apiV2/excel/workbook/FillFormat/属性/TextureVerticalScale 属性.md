---
title: "TextureVerticalScale 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/TextureVerticalScale%20%E5%B1%9E%E6%80%A7.html
---
# TextureVerticalScale (属性)

返回指定填充的纹理垂直缩放比例因子。可读/写 **Single** 类型。

## 示例

```javascript
/*本示例将第一张工作表中第一个形状填充的纹理垂直缩放比例因子设置为0.5。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    fillFormat.TextureVerticalScale = 0.5
}
```

```javascript
/*本示例显示活动工作表中第二个形状填充的纹理垂直缩放比例因子。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    console.log(fillFormat.TextureVerticalScale)
}
```
