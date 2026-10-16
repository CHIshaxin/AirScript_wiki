---
title: "PresetTexture (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/PresetTexture%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CXEvTt4Y
---

# PresetTexture (属性)

返回指定填充的预设纹理。**MsoPresetTexture**，只读。

## 说明

使用 **PresetTextured**方法可设置填充的预设纹理。

## 示例

```javascript
/*本示例显示第一张工作表中第一个形状填充的预设纹理是否为水滴纹理。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    console.log(fillFormat.PresetTexture == msoTextureWaterDroplets)
}
```

```javascript
/*本示例判断活动工作表中第二个形状的预设纹理。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    if (fillFormat.PresetTexture == msoPresetTextureMixed) {
        console.log("未使用纹理")
    } else {
        console.log("其它纹理")
    }
}
```
