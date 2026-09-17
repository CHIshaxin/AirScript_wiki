---
title: "Weight 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Borders/%E5%B1%9E%E6%80%A7/Weight%20%E5%B1%9E%E6%80%A7.html
---
# Weight (属性)

返回或设置一个 **XlBorderWeight**值，它代表边框的粗细。

## 示例

```javascript
/*本示例将 E6 单元格边框的粗细设置为 xlThick 。*/
function test() {
    Range("E6").Borders.Weight = xlThick
}
```

```javascript
/*本示例显示 D3 单元格边框的粗细是否为 xlThin 。*/
function test() {
    console.log(Range("D3").Borders.Weight == xlThin)
}
```
