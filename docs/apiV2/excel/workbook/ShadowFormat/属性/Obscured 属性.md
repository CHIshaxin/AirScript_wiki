---
title: "Obscured 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/Obscured%20%E5%B1%9E%E6%80%A7.html
---
# Obscured (属性)

如果指定形状的阴影是填充的，并且阴影被形状所遮盖（即便该形状没有填充），那么该值为 **True**。如果阴影没有填充，并且当形状没有填充时，可透过形状看到阴影的轮廓，那么该值为 **False**。**MsoTriState** 类型，可读写。

## 说明

| **MsoTriState**可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** |
| **msoFalse**阴影没有填充；如果形状也没有填充，透过形状可看见阴影的轮廓。 |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue**指定形状的阴影有填充并被该形状所遮蔽，即使形状也没有填充。 |

## 示例

```javascript
/*本示例显示第一张工作表中的形状 3 阴影是否被遮盖。*/
function test() {
    let shadow = Worksheets.Item(1).Shapes.Item(3).Shadow
    console.log(shadow.Obscured == msoTrue)
}
```

```javascript
/*本示例显示活动工作表中第二个形状阴影是否被遮盖。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(2).Shadow
    console.log(shadow.Obscured == msoTrue)
}
```
