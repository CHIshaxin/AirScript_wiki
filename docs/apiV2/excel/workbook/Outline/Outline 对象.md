---
title: "Outline 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Outline/Outline%20%E5%AF%B9%E8%B1%A1.html
---
# Outline (对象)

代表工作表上的分级显示。

## 示例

```javascript
/*下例将 Sheet4 上的分级显示设置为只显示第一级。*/
function test() {
    Worksheets("Sheet4").Outline.ShowLevels(1)
}
```

```javascript
/*此示例显示活动工作表分级显示是否使用自动样式。*/
function test() {
    console.log(ActiveSheet.Outline.AutomaticStyles)
}
```
