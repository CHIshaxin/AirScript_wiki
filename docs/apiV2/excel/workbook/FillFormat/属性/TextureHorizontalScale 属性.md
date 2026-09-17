---
title: "TextureHorizontalScale 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/TextureHorizontalScale%20%E5%B1%9E%E6%80%A7.html
---
# TextureHorizontalScale (属性)

返回或设置纹理填充的水平缩放比例因子。可读/写 **Single** 类型。

## 示例

```javascript
/*本示例将第一张工作表中第一个形状纹理填充的水平缩放比例因子设置为0.4。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    fillFormat.TextureHorizontalScale = 0.4
}
```

```javascript
/*本示例显示活动工作表中第二个形状纹理填充的水平缩放比例因子。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    console.log(fillFormat.TextureHorizontalScale)
}
```
