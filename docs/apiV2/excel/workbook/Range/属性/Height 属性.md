---
title: "Height 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
---
# Height (属性)

返回一个 Double 值，该值代表区域的高度（以磅为单位）。此为只读属性。

## 示例

```javascript
/*本示例使用 RowHeight 设置 A1 单元格的行高为 15，并显示 A1 单元格的行高*/
function test() {
    Range("A1").RowHeight = 15
    console.log(Range("A1").Height)
}
```

```javascript
/*本示例显示 A1 单元格 Height 的数据类型*/
function test() {
    console.log(typeof Range("A1").Height)
}
```
