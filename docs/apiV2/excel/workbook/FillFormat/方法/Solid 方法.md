---
title: "Solid 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E6%96%B9%E6%B3%95/Solid%20%E6%96%B9%E6%B3%95.html
---
# Solid (方法)

将指定填充设为统一的颜色。使用此方法可将渐变、纹理、图案或背景填充转换为纯色填充。

## 示例

```javascript
/*此示例将第一张工作表中的所有填充转换为统一的红色填充。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        let fill = shapes.Item(i).Fill
        fill.Solid()
        fill.ForeColor.RGB = RGB(255, 0, 0)
    }
}
```

```javascript
/*本示例将活动工作表中第二个形状的填充格式设置为纯色填充。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    fillFormat.Solid()
}
```
