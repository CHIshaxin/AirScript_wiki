---
title: "TextureAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/TextureAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BZDe_iO1
---

# TextureAlignment (属性)

返回或设置用于平铺纹理填充的对齐方式（坐标网格的原点）。可读/写。

## 示例

```javascript
/*本示例将第一张工作表中第一个形状平铺纹理填充的对齐方式设置为左对齐。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    fillFormat.TextureAlignment = msoTextureLeft
}
```

```javascript
/*本示例显示活动工作表中第二个形状平铺纹理填充的对齐方式是否为从上到左对齐。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    console.log(fillFormat.TextureAlignment == msoTextureTopLeft)
}
```
