---
title: "Style 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Style%20%E5%B1%9E%E6%80%A7.html
---
# Style (属性)

返回或设置一个包含 **[Style](./../../Style/Style 对象.html)** 对象的 **Variant** 值，它代表指定区域的样式。

## 示例

```javascript
/*此示例对工作表 sheet1 中的单元格 A1 应用“常规”样式。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").Style = "Normal"
}
```

```javascript
/*如果工作表 sheet1 中的单元格 B4 当前应用的是“常规”样式，此示例将对其应用“百分比”样式。*/
function test() {
    if (Worksheets.Item("Sheet1").Range("B4").Style == "Normal") {
        Worksheets.Item("Sheet1").Range("B4").Style = "Percent"
    }
}
```
