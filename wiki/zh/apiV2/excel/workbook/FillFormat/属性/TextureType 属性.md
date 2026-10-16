---
title: "TextureType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/TextureType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BsujHrXR
---

# TextureType (属性)

返回指定填充的纹理类型。**MsoTextureType**，只读。

## 说明

使用 **UserTextured**方法可设置填充的纹理类型。

## 示例

```javascript
/*此示例以第一个形状中使用的填充样式对第二个形状的填充格式进行设置。*/
function test() {
    let fillFormat1 = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    if (fillFormat1.Type == msoFillTextured) {
        let fillFormat2 = Application.Worksheets.Item(1).Shapes.Item(2).Fill
        fillFormat2.Visible = true
        if (fillFormat1.TextureType == msoTexturePreset) {
            fillFormat2.PresetTextured(fillFormat1.PresetTexture)
        } else {
            fillFormat2.UserTextured(fillFormat1.TextureName)
        }
    }
}
```

```javascript
/*本示例显示活动工作表中第一个形状填充的纹理类型是否为预设纹理类型。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(1).Fill
    console.log(fillFormat.TextureType == msoTexturePreset)
}
```
