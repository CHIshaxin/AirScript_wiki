---
title: "Style (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/Style%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BVSrp59B
---

# Style (属性)

返回或设置指定底纹的样式。可读/写 **MsoShadowStyle** 类型。

## 说明

| MsoShadowStyle 可以是下列 **MsoShadowStyle** 常量之一。 |
| --- |
| **msoShadowStyleInnerShadow** |
| **msoShadowStyleMixed** |
| **msoShadowStyleOuterShadow** |

## 示例

```javascript
/*本示例显示第一张工作表中第二个形状阴影的样式是否为msoShadowStyleMixed。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(2).Shadow
    console.log(shadow.Style == msoShadowStyleMixed)
}
```

```javascript
/*本示例设置活动工作表中第一个形状阴影的样式。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(1).Shadow
    shadow.Style = msoShadowStyleInnerShadow
}
```
