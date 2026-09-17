---
title: "LinearGradient 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LinearGradient/LinearGradient%20%E5%AF%B9%E8%B1%A1.html
---
# LinearGradient (对象)

**[LinearGradient](./../LinearGradient/LinearGradient 对象.html)** 对象沿特定角度以线性方式在一系列颜色间转换。

## 说明

当使用 **[LinearGradient](./../LinearGradient/LinearGradient 对象.html)** 对象时，应考虑以下几点：

- 试图访问不具有现有渐变填充的 **[Interior](./../Interior/Interior 对象.html)** 对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意 `Interior.Pattern` 属性。
- 如果将 Interior.Pattern从渐变类型更改为非渐变类型，Gradient 对象将采用默认值。

## 示例

```javascript
/*下例将活动工作表上区域 B1:B10 中线性渐变的填充角度设置为100。*/
function test() {
    let linearGradient = ActiveSheet.Range("B1:B10").Interior.Gradient
    linearGradient.Degree = 100
}
```

```javascript
/*下例显示第一张工作表上区域 B1:B10 中线性渐变的颜色停止点数量。*/
function test() {
    let linearGradient = Worksheets.Item(1).Range("B1:B10").Interior.Gradient
    console.log(linearGradient.ColorStops.Count)
}
```
