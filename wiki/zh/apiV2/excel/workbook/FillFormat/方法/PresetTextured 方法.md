---
title: "PresetTextured (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E6%96%B9%E6%B3%95/PresetTextured%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B2QsHeyB
---

# PresetTextured (方法)

将指定填充格式设为预设纹理。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| PresetTexture | Kso.MsoPresetTexture | 必填 | 要应用的纹理类型。 |

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
/*本示例将活动工作表中第二个形状的预设纹理设置成棕色大理石纹理。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    fillFormat.PresetTextured(msoTextureBrownMarble)
}
```
