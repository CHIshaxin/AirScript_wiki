---
title: "Transparency (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/%E5%B1%9E%E6%80%A7/Transparency%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: QsVsCMPG
---

# Transparency (属性)

返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。**Double** 型，可读写。

## 说明

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

## 示例

```javascript
/*本示例将第一张工作表中第二个形状线条的透明度设置为0.3。*/
function test() {
    let line = Worksheets.Item(1).Shapes.Item(2).Line
    line.Transparency = 0.3
}
```

```javascript
/*本示例显示活动工作表的第一个形状线条的透明度。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    console.log(line.Transparency)
}
```
