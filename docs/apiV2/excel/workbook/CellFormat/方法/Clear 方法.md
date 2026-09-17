---
title: "Clear 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
---
# Clear (方法)

清除 **FindFormat** 和 **ReplaceFormat** 属性中的条件集。

## 示例

```javascript
/*本示例清除 FindFormat 中设置的条件。*/
function test() {
    Application.FindFormat.Clear()
}
```

```javascript
/*本示例清除 ReplaceFormat 中设置的条件。*/
function test() {
    Application.ReplaceFormat.Clear()
}
```
