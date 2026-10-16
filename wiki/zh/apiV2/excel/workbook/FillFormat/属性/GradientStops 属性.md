---
title: "GradientStops (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/GradientStops%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 4Lmul1aY
---

# GradientStops (属性)

返回渐变填充的终点。只读。

## 示例

```javascript
/*本示例将活动工作表中第一个形状设置成具有双色渐变填充效果。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(1).Fill
    fillFormat.TwoColorGradient(msoGradientHorizontal, 1)
    fillFormat.GradientStops.Insert2(RGB(255, 0, 0), 0.5, 0, 1)
}
```

```javascript
/*本示例显示第一张工作表中第二个形状 GradientStops 集合中的项数。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(2).Fill
    console.log(fillFormat.GradientStops.Count)
}
```
