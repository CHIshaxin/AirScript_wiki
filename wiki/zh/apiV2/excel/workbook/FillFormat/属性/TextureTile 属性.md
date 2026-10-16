---
title: "TextureTile (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/TextureTile%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DB71d-MX
---

# TextureTile (属性)

返回指定填充的纹理平铺样式。可读/写 **MsoTriState** 类型。

## 说明

| MsoTriState 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** |
| **msoFalse** |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue** |

## 示例

```javascript
/*本示例显示活动工作表中第一个形状填充的纹理是否为平铺。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(1).Fill
    console.log(fillFormat.TextureTile == msoTrue)
}
```

```javascript
/*本示例将第一张工作表中第二个形状填充的纹理设置为拉伸方式。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(2).Fill
    fillFormat.TextureTile = msoFalse
}
```
